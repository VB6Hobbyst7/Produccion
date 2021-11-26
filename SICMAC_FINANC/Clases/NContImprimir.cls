VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NContImprimir"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A7EE84101F4"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Dim vsBaseComunes As String
Dim vsBasePesonas As String
Dim vsAgeNom As String
Dim vsEmpNom As String
Dim vsFecSis As String
Dim lsEmprLogo As String
'Type tBilletajes
'    lsBilletes As Currency
'    lnMontoBil As Currency
'    lsMonedas As Currency
'    lnMontoMon As Currency
'End Type

Public Event BarraShow(pnMax)
Public Event BarraProgress(value, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
Public Event BarraClose()

'RIRO20140630 ERS017 nTasa, nDias
Public Function ImprimeAsientoContable(ByVal psMovNro As String, ByVal pnLinPage As Integer, ByVal pnColPage As Integer, Optional ByVal psTit1 As String, Optional ByVal psCabe1 As String, Optional ByVal psPiePag As String = "9", Optional ByVal psOtraFirma As String = "", Optional psNomCmact As String = "", Optional pbHistoCtaCont As Boolean = False, Optional nTasa As Double = 0, Optional nDias As Integer = 0) As String
    On Error GoTo ImprimeAsientoContableErr
    Dim sSQL As String
    Dim rs       As New ADODB.Recordset
    Dim rsDoc       As New ADODB.Recordset
    Dim nTot     As Currency
    Dim nTotH    As Currency
    Dim sDoc     As String
    Dim sTexto   As String
    Dim sAsiento As String
    Dim nLi      As Integer
    Dim sCabecera As String
    Dim lnCambioFijo As Currency
    Dim lnCambioVenta As Currency
    Dim lnCambioCompra As Currency
    Dim lnCambioVentaSBS As Currency
    Dim lnCambioCompraSBS As Currency
    Dim CON As String
    Dim COFF As String
    Dim lsMovBilletaje As String
    
    Dim lsDesObjeto As String
    Dim lsFecha    As String
    Dim oConect As DConecta
    Dim lsGlosa As String
    Dim lsMovNro As String
    Dim lsOpeCod As String
    Dim sDestino As String
    Dim sOrigen As String
    Dim oDGeneral As DGeneral
    Dim BON As String
    Dim BOFF As String
    Dim oTipoCamb As nTipoCambio
    
    Set oDGeneral = New DGeneral
    Set oTipoCamb = New nTipoCambio
    
    Set oConect = New DConecta
    
    If oConect.AbreConexion = False Then Exit Function
    CON = PrnSet("C+")
    COFF = PrnSet("C-")
    BON = PrnSet("B+")
    BOFF = PrnSet("B-")
    
    sCabecera = ""
    sAsiento = ""
    nLi = 0
    If psMovNro = "" Then
       Exit Function
    End If
       lsFecha = GetFechaMov(psMovNro, True)
       sSQL = "SELECT * FROM Mov WHERE cMovNro = '" & psMovNro & "'"
       Set rs = oConect.CargaRecordSet(sSQL)
       If rs.EOF Then
          Exit Function
       End If
       lsGlosa = Replace(rs!cMovdesc, Chr(13) & oImpresora.gPrnSaltoLinea, " ")
       lsMovNro = psMovNro
       lsOpeCod = rs!cOpeCod
       If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Or Left(lsOpeCod, 4) = Left(gOpeME, 4) Or rs!cOpeCod = gContAjusteTipoCambio Then
         lnCambioFijo = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCFijoMes)
         lnCambioVenta = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCVenta)
         lnCambioCompra = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCCompra)
         
         lnCambioVentaSBS = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), 7)
         lnCambioCompraSBS = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCPonderado)
       End If
       
       If rs!nMovFlag = gMovFlagDeExtorno Then
            psTit1 = " A S I E N T O  C O N T A B L E   D E   E X T O R N O"
       ElseIf rs!nMovFlag = gMovFlagEliminado Then
            psTit1 = " **** ASIENTO ELIMINADO NO VALIDO PARA CONTABILIDAD **** "
       End If
       'ddd
       If Left(lsOpeCod, 3) = "622" Or Left(lsOpeCod, 3) = "632" Then
            sAsiento = sAsiento & ImpreCabAsiento(pnColPage, CDate(lsFecha), gsNomCmac, lsOpeCod, lsMovNro, psTit1, False, , , nLi) & oImpresora.gPrnSaltoLinea
       Else
            sAsiento = sAsiento & ImpreCabAsiento(pnColPage, CDate(lsFecha), gsNomCmac, lsOpeCod, lsMovNro, psTit1, True, , , nLi) & oImpresora.gPrnSaltoLinea
       End If
       If psCabe1 <> "" Then
          sAsiento = sAsiento & psCabe1 & oImpresora.gPrnSaltoLinea
          nLi = nLi + 4
       Else
           'BUSCA LAS PERSONA INVOLUCRADAS EN EL MOVIMIENTO
            sSQL = " SELECT  DISTINCT M.CMOVNRO,  " _
                    & "         CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                    & "                 WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                    & "         P.cPersCod, " _
                    & "         Convert(Char(40),P.CPERSNOMBRE) as cPersNombre, " _
                    & "         MOI.CPersCod " _
                    & " FROM    MOV M JOIN MOVCTA MC    ON M.NMOVNRO = MC.NMOVNRO " _
                    & "         JOIN MOVGasto MOI         ON MOI.NMOVNRO = M.NMOVNRO " _
                    & "         JOIN PERSONA P          ON P.CPERSCOD = MOI.CPERSCOD " _
                    & " WHERE   M.cMovNro = '" & lsMovNro & "' "
              
            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = ""
            Do While Not rs.EOF
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & PstaNombre(rs!cpersNombre)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & PstaNombre(rs!cpersNombre)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat("Persona(s) Origen  ", 25) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat("Persona(s) Destino ", 25) & " :" & sDestino & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
            sSQL = " SELECT  DISTINCT M.CMOVNRO, " _
                & "         CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                 WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "         P.cPersCod , CI.cCtaIFCod , " _
                & "         Convert(Char(40),CTPOI.cConsDescripcion ) as Entidad, " _
                & "         Convert(Char(40),P.CPERSNOMBRE) as cPersNombre, " _
                & "         Convert(char(50),C.cConsDescripcion + ' ' + CI.cCtaIFDesc) as DescCtaIF, " _
                & "         MO.COBJETOCOD " _
                & " FROM    MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "         JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "         JOIN MOVOBJIF MOI       ON MOI.NMOVNRO = MO.NMOVNRO AND MOI.NMOVITEM = MO.NMOVITEM " _
                & "         JOIN CTAIF CI           ON CI.CCTAIFCOD = MOI.CCTAIFCOD AND MOI.CPERSCOD = CI.CPERSCOD AND MOI.CIFTPO = CI.CIFTPO  " _
                & "         JOIN CONSTANTE C        ON C.nConsValor = SUBSTRING(CI.CCTAIFCOD,2,1) AND C.nCONSCOD LIKE '" & gCGTipoCtaIF & "' " _
                & "         JOIN INSTITUCIONFINANC I    ON I.CPERSCOD = CI.CPERSCOD AND I.CIFTPO = CI.CIFTPO " _
                & "         JOIN PERSONA P          ON P.CPERSCOD = I.CPERSCOD " _
                & "         JOIN CONSTANTE CTPOI        ON CTPOI.nConsValor = I.cIFTpo AND CTPOI.nCONSCOD LIKE '" & gCGTipoIF & "' " _
                & " WHERE M.cMovNro = '" & lsMovNro & "' AND MO.cObjetoCod ='" & Format(ObjEntidadesFinancieras, "00") & "'"
            
            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cpersNombre) & " " & Trim(rs!DescCtaIF)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cpersNombre) & " " & Trim(rs!DescCtaIF)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat("Cuenta(s) Origen", 20) & " :" + PrnSet("C+") + sOrigen + PrnSet("C-") + oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat("Cuenta(s) Destino", 20) & " :" + PrnSet("C+") + sDestino + PrnSet("C-") + oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
            'RIRO20140630 ERS017 ****
            If nTasa <> 0 And nDias <> 0 Then
                sAsiento = sAsiento & ImpreFormat("Plazo Vencimiento(s) ", 20) & " : " & nDias & " dias(s)          Tasa: " & nTasa & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            'END RIRO ***************
            
            sSQL = "    SELECT  DISTINCT M.CMOVNRO, " _
                & "             CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                     WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "             MOA.cAreaCod + MOA.CAGECOD AS cAreaCod, A.cAreaDescripcion + ' ' + ISNULL(AG.CAGEDESCRIPCION,'' )  as cAreaDescripcion , MO.COBJETOCOD, O.cObjetoDesc " _
                & "     FROM    MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "             JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "             JOIN MovObjAreaAgencia MOA         ON MOA.NMOVNRO = MO.NMOVNRO AND MOA.NMOVITEM = MO.NMOVITEM " _
                & "             JOIN AREAS A            ON A.cAreaCod = MOA.cAreaCod LEFT JOIN AGENCIAS AG ON AG.CAGECOD = MOA.CAGECOD " _
                & "             JOIN OBJETO O ON O.cObjetoCod  = MO.cObjetoCod " _
                & "     WHERE   M.cMovNro = '" & lsMovNro & "' AND Mo.cObjetoCod IN ('" & Format(ObjCMACAgenciaArea, "00") & "','" & Format(ObjCMACArea, "00") & "')"

            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                lsDesObjeto = Trim(rs!cObjetoDesc)
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!cAreaDescripcion)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!cAreaDescripcion)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Origen ", 20) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Destino ", 20) & " :" & sDestino & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If

            sSQL = " SELECT    DISTINCT M.CMOVNRO, " _
                & "             CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                     WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "             MO.COBJETOCOD, O.cObjetoDesc " _
                & "   FROM      MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "             JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "             JOIN OBJETO O ON O.cObjetoCod  = MO.cObjetoCod " _
                & "   WHERE     M.cMovNro='" & lsMovNro & "' " _
                & "             AND MO.cObjetoCod NOT IN ('" & Format(ObjCMACAgencias, "00") & "','" & Format(ObjCMACAgenciaArea, "00") & "','" & Format(ObjDescomEfectivo, "00") & "','" & Format(ObjEntidadesFinancieras, "00") & "','" & Format(ObjBienesServicios, "00") & "','" & Format(ObjPersona, "00") & "')"
            
            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                lsDesObjeto = "OBJETOS"
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cObjetoCod) & "-" & Trim(rs!cObjetoDesc)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cObjetoCod) & "-" & Trim(rs!cObjetoDesc)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Origen ", 20) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Destino", 20) & " :" & sDestino & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
        End If ' del pscab1

       'If lsOpeCod <> "500520" And lsOpeCod <> "500510" Then '** Juez 20120625
       If lsOpeCod <> "500520" And lsOpeCod <> "500510" And lsOpeCod <> 900022 And lsOpeCod <> 900023 Then '** Juez 20120625 -JUCS 20171104
        'Tipo de Cambio
         sSQL = "SELECT nMovTpoCambio, m.cOpeCod FROM MovTpoCambio MC JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE M.cMovNro = '" & lsMovNro & "'"
         Set rs = oConect.CargaRecordSet(sSQL)
         If Not rs.EOF Then
             If rs!cOpeCod = gContAjusteTipoCambio Then
             Else
                sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "       T.C.C SBS : " & Format(lnCambioCompraSBS, "##,###,##0.000") & "     T.C.V SBS : " & Format(lnCambioVentaSBS, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
                'sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "      T.C.Mercado: " & Format(rs!nMovTpoCambio, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
             End If
             nLi = nLi + 1
         Else
             If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Then
                 sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "     T.C.Compra : " & Format(lnCambioCompra, "##,###,##0.000") & "     T.C.Venta : " & Format(lnCambioVenta, "##,###,##0.000") & "     T.C.C SBS : " & Format(lnCambioCompraSBS, "##,###,##0.000") & "     T.C.V SBS : " & Format(lnCambioVentaSBS, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
                 nLi = nLi + 1
             End If
         End If
        End If
        '** Juez 20120625 *************************************************
        If lsOpeCod = "500520" Then
            sAsiento = sAsiento & "  Fecha: " & lsFecha & oImpresora.gPrnSaltoLinea '& Space(132 - 45) & "Usuario: " & FillText(gsCodUser, 5, " ") & oImpresora.gPrnSaltoLinea
            'sAsiento = sAsiento & Space(132 - 50) & "Area: " & FillText(gsCodUser, 30, " ") & oImpresora.gPrnSaltoLinea
        End If
        If lsOpeCod = "500510" Then
            Dim oConting As DContingencia
            Dim rsConting As ADODB.Recordset
            Set oConting = New DContingencia
            Set rsConting = oConting.ObtieneUltimoInformeTecConting(oConting.ObtieneNroRegistroxnMovNroIT(lsMovNro))
            If Not rsConting.EOF Then 'Juez 20120921
                sAsiento = sAsiento & "  Informe N*: " & rsConting!cNroInforme & oImpresora.gPrnSaltoLinea
                sAsiento = sAsiento & "  Fecha de Informe: " & Format(rsConting!dFechaRegInfTec, "dd/mm/yyyy") & oImpresora.gPrnSaltoLinea
            Else
                sAsiento = sAsiento & "  Informe N*: " & oImpresora.gPrnSaltoLinea
                sAsiento = sAsiento & "  Fecha de Informe: " & oImpresora.gPrnSaltoLinea
            End If
        End If
        'Mejoras  TI-ERS 002-2017 JUCS
       If lsOpeCod = "900022" Or lsOpeCod = "900023" Then
          'Tipo de Cambio
          sSQL = "exec stp_sel_ObtieneCambioTCCompraVenta  '" & lsMovNro & "' "
          Set rs = oConect.CargaRecordSet(sSQL)
          If Not rs.EOF Then
          sAsiento = sAsiento + "  T.C.Fijo : " & Format(rs!nValFijo, "##,###,##0.000") & "    T.C.Compra : " & Format(rs!nCompra, "##,###,##0.000") & "      T.C.Venta : " & Format(rs!nVenta, "##,###,##0.000") & "     T.C.C SBS : " & Format(rs!nValPond, "##,###,##0.000") & "     T.C.V SBS : " & Format(rs!nValPondVenta, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
       Else
          MsgBox "Esta Operacion no cuenta con algun Tipo de cambio registrado", vbInformation, "Informacion"
       End If
       End If
        '**End JUCS*****************************************************************************
        '** End Juez ******************************************************
        nLi = nLi + 1
       'Datos de Documentos
       sSQL = "SELECT b.cDocAbrev, a.cDocNro, a.dDocFecha FROM MovDoc a join mov m on m.nmovnro =a.nmovnro  , " & vsBaseComunes & "Documento b " _
            & "WHERE m.cMovNro = '" & lsMovNro & "' and b.nDocTpo = a.nDocTpo "

        Set rsDoc = oConect.CargaRecordSet(sSQL)
        If Not rsDoc.EOF Then
          sDoc = ""
          nLi = 1
          Do While Not rsDoc.EOF
             sDoc = sDoc & rsDoc!cDocAbrev & "_" & rsDoc!cDocNro & " " & rsDoc!dDocFecha & "  "
             If nLi Mod 4 = 0 Then
                sDoc = sDoc + oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
             End If
             rsDoc.MoveNext
          Loop
          sAsiento = sAsiento + ImpreFormat("Documentos", 20) & " :" + PrnSet("C+") + sDoc + PrnSet("C-") + oImpresora.gPrnSaltoLinea
          nLi = nLi + 1
        End If
        rsDoc.Close
       'Glosa
        If lsOpeCod <> "500520" And lsOpeCod <> "500510" Then '** Juez 20120625
            sAsiento = sAsiento & ImpreGlosa(lsGlosa, pnColPage, ImpreFormat("Glosa       ", 20) & " :", , , nLi) + oImpresora.gPrnSaltoLinea
        Else
            sAsiento = sAsiento & "  Glosa: " & lsGlosa & oImpresora.gPrnSaltoLinea
        End If
       nLi = nLi + 1
       lsMovBilletaje = ""
       lsMovBilletaje = ImprimeBilletaje(Mid(lsOpeCod, 3, 1), lsMovNro, pnColPage, , nLi)
       If lsMovBilletaje <> "" Then
            sAsiento = sAsiento + lsMovBilletaje
       End If
    'Cabecera 2 de Asiento
    sCabecera = sCabecera + CON
    sCabecera = sCabecera + "====================================================================================================================================" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "Itm C U E N T A   C O N T A B L E                                                                           DEBE            HABER" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "    Codigo              Descripcion" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "------------------------------------------------------------------------------------------------------------------------------------" & COFF & oImpresora.gPrnSaltoLinea
    'MovCta para MN
    
    'JACA 20111227 Modificado para obtener cCtaContDesc Historica**************************
    sSQL = "SELECT mc.nMovItem, mc.cCtaContCod," + IIf(pbHistoCtaCont = False, "ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),'')", "ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),'')") + " cCtaContDesc, mc.nMovImporte " _
         & "FROM   MovCta mc JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE  M.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
'    sSQL = "SELECT mc.nMovItem, mc.cCtaContCod," + IIf(pbHistoCtaCont = False, "ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),'')", "ISNULL(dbo.Histo_GetCtaContDesc(mc.cCtaContCod,2,1),'')") + " cCtaContDesc, mc.nMovImporte " _
'         & "FROM   MovCta mc JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE  M.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
    'JACA END******************************************************************
    
    Set rs = oConect.CargaRecordSet(sSQL)
    If rs.EOF Then
       'MsgBox "No existen datos de Movimiento ", vbInformation, "Error"
       Exit Function
    End If
    nTot = 0
    nTotH = 0
    sAsiento = sAsiento + sCabecera
    nLi = nLi + 4
    sAsiento = sAsiento + BON & " EN MONEDA NACIONAL " + oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento + "--------------------" & BOFF + oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento & oImpresora.gPrnSaltoLinea
    Do While Not rs.EOF
       DoEvents
       If nLi > pnLinPage - 6 Then
           nLi = 5
           sAsiento = sAsiento + oImpresora.gPrnSaltoPagina & sCabecera
       End If
       sAsiento = sAsiento + CON & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " & Justifica(rs!cCtaContDesc, 71) & " " _
            & Right(Space(16) & IIf(rs!nMovImporte > 0, Format(rs!nMovImporte, "##,###,##0.00"), ""), 16) & " " _
            & Right(Space(16) & IIf(rs!nMovImporte < 0, Format(rs!nMovImporte * -1, "##,###,##0.00"), ""), 16) _
            & COFF & oImpresora.gPrnSaltoLinea
       If rs!nMovImporte > 0 Then
          nTot = nTot + Val(rs!nMovImporte)
       Else
          nTotH = nTotH + Val(rs!nMovImporte) * -1
       End If
       nLi = nLi + 1
       rs.MoveNext
    Loop
    sAsiento = sAsiento + CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF & oImpresora.gPrnSaltoLinea
    nLi = nLi + 2

    'JACA 20111227 Modificado para obtener cCtaContDesc Historica**************************
    sSQL = "SELECT mc.nMovItem, mc.cCtaContCod," + IIf(pbHistoCtaCont = False, " LTRIM(RTRIM(ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),'')))", " LTRIM(RTRIM(ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),'')))") + " cCtaContDesc, me.nmovMEimporte " _
         & "FROM   mov m join MovCta mc on mc.nmovnro=m.nmovnro JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
         & "WHERE  m.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
'    sSQL = "SELECT mc.nMovItem, mc.cCtaContCod," + IIf(pbHistoCtaCont = False, " LTRIM(RTRIM(ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),'')))", " LTRIM(RTRIM(ISNULL(dbo.Histo_GetCtaContDesc(mc.cCtaContCod,2,1),'')))") + " cCtaContDesc, me.nmovMEimporte " _
'         & "FROM   mov m join MovCta mc on mc.nmovnro=m.nmovnro JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'         & "WHERE  m.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
    Set rs = oConect.CargaRecordSet(sSQL)
    'JACA END**************************************************************************
    sTexto = ""
    nTot = 0
    nTotH = 0
    If Not rs.EOF Then
       sAsiento = sAsiento + BON & " EN MONEDA EXTRANJERA " & oImpresora.gPrnSaltoLinea
       sAsiento = sAsiento + "----------------------" & BOFF & oImpresora.gPrnSaltoLinea
       nLi = nLi + 2
    End If
    Do While Not rs.EOF
       DoEvents
       If nLi > pnLinPage - 6 Then
          nLi = 5
          sTexto = sTexto + oImpresora.gPrnSaltoPagina
          sTexto = sTexto + sCabecera
       End If
       sTexto = sTexto + CON & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " _
              & Justifica(rs!cCtaContDesc, 71) & " " _
              & IIf(rs!nMovMEImporte > 0, PrnVal(rs!nMovMEImporte, 16, 2), Space(16)) & " " _
              & IIf(rs!nMovMEImporte < 0, PrnVal(rs!nMovMEImporte * -1, 16, 2), Space(16)) _
              & COFF & oImpresora.gPrnSaltoLinea
       If rs!nMovMEImporte > 0 Then
          nTot = nTot + Val(rs!nMovMEImporte)
       Else
          nTotH = nTotH + Val(rs!nMovMEImporte) * -1
       End If
       nLi = nLi + 1
       rs.MoveNext
    Loop
    If nTot > 0 Or nTotH > 0 Then
       sAsiento = sAsiento & sTexto
       sAsiento = sAsiento + CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
       sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF + oImpresora.gPrnSaltoLinea
       nLi = nLi + 3
    End If
    If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
    
    If lsOpeCod <> 421113 Then
        sAsiento = sAsiento + CON & "====================================================================================================================================" & COFF & oImpresora.gPrnSaltoLinea
        
        sAsiento = sAsiento + ImprePiePag(pnColPage, psPiePag, psOtraFirma)
    End If
    sAsiento = sAsiento + "   " + oImpresora.gPrnSaltoLinea

    Set oDGeneral = Nothing
    Set oConect = Nothing
    ImprimeAsientoContable = sAsiento
    
    Exit Function
ImprimeAsientoContableErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeAsientoContable Method")
End Function

Public Function ImprimeAsientoNoContable(ByVal psMovNro As String, ByVal pnLinPage As Integer, ByVal pnColPage As Integer, Optional ByVal psTit1 As String, Optional ByVal psCabe1 As String, Optional ByVal psPiePag As String = "9", Optional ByVal psOtraFirma As String = "", Optional psNomCmact As String = "") As String
    On Error GoTo ImprimeAsientoNoContableErr
    Dim sSQL As String
    Dim rs       As New ADODB.Recordset
    Dim rsDoc       As New ADODB.Recordset
    Dim nTot     As Currency
    Dim nTotH    As Currency
    Dim sDoc     As String
    Dim sTexto   As String
    Dim sAsiento As String
    Dim nLi      As Integer
    Dim sCabecera As String
    Dim lnCambioFijo As Currency
    Dim lnCambioVenta As Currency
    Dim lnCambioCompra As Currency
    Dim CON As String
    Dim COFF As String
    Dim lsMovBilletaje As String
    
    Dim lsDesObjeto As String
    Dim lsFecha    As String
    Dim oConect As DConecta
    Dim lsGlosa As String
    Dim lsMovNro As String
    Dim lsOpeCod As String
    Dim sDestino As String
    Dim sOrigen As String
    Dim oDGeneral As DGeneral
    Dim BON As String
    Dim BOFF As String
    Dim oTipoCamb As nTipoCambio
    
    Set oDGeneral = New DGeneral
    Set oTipoCamb = New nTipoCambio
    
    Set oConect = New DConecta
    
    If oConect.AbreConexion = False Then Exit Function
    CON = PrnSet("C+")
    COFF = PrnSet("C-")
    BON = PrnSet("B+")
    BOFF = PrnSet("B-")
    
    sCabecera = ""
    sAsiento = ""
    nLi = 0
    If psMovNro = "" Then
       Exit Function
    End If
       lsFecha = GetFechaMov(psMovNro, True)
       sSQL = "SELECT * FROM Mov WHERE cMovNro = '" & psMovNro & "'"
       Set rs = oConect.CargaRecordSet(sSQL)
       If rs.EOF Then
          Exit Function
       End If
       lsGlosa = Replace(rs!cMovdesc, Chr(13) & oImpresora.gPrnSaltoLinea, " ")
       lsMovNro = psMovNro
       lsOpeCod = rs!cOpeCod
       If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Or Left(lsOpeCod, 4) = Left(gOpeME, 4) Or rs!cOpeCod = gContAjusteTipoCambio Then
         lnCambioFijo = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCFijoMes)
         lnCambioVenta = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCVenta)
         lnCambioCompra = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCCompra)
       End If
       
       If rs!nMovFlag = gMovFlagDeExtorno Then
            psTit1 = " A S I E N T O  C O N T A B L E   D E   E X T O R N O"
       ElseIf rs!nMovFlag = gMovFlagEliminado Then
            psTit1 = " **** ASIENTO ELIMINADO NO VALIDO PARA CONTABILIDAD **** "
       End If
       sAsiento = sAsiento & ImpreCabAsiento(pnColPage, CDate(lsFecha), gsNomCmac, lsOpeCod, lsMovNro, psTit1, True, , , nLi) & oImpresora.gPrnSaltoLinea
       If psCabe1 <> "" Then
          sAsiento = sAsiento & psCabe1 & oImpresora.gPrnSaltoLinea
          nLi = nLi + 4
       Else
           'BUSCA LAS PERSONA INVOLUCRADAS EN EL MOVIMIENTO
            sSQL = " SELECT  DISTINCT M.CMOVNRO,  " _
                    & "         CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                    & "                 WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                    & "         P.cPersCod, " _
                    & "         Convert(Char(40),P.CPERSNOMBRE) as cPersNombre, " _
                    & "         MOI.CPersCod " _
                    & " FROM    MOV M JOIN MOVCTA MC    ON M.NMOVNRO = MC.NMOVNRO " _
                    & "         JOIN MOVGasto MOI         ON MOI.NMOVNRO = M.NMOVNRO " _
                    & "         JOIN PERSONA P          ON P.CPERSCOD = MOI.CPERSCOD " _
                    & " WHERE   M.cMovNro = '" & lsMovNro & "' "
              
            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = ""
            Do While Not rs.EOF
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & PstaNombre(rs!cpersNombre)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & PstaNombre(rs!cpersNombre)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat("Persona(s) Origen  ", 25) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat("Persona(s) Destino ", 25) & " :" & sDestino & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
            sSQL = " SELECT  DISTINCT M.CMOVNRO, " _
                & "         CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                 WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "         P.cPersCod , CI.cCtaIFCod , " _
                & "         Convert(Char(40),CTPOI.cConsDescripcion ) as Entidad, " _
                & "         Convert(Char(40),P.CPERSNOMBRE) as cPersNombre, " _
                & "         Convert(char(50),C.cConsDescripcion + ' ' + CI.cCtaIFDesc) as DescCtaIF, " _
                & "         MO.COBJETOCOD " _
                & " FROM    MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "         JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "         JOIN MOVOBJIF MOI       ON MOI.NMOVNRO = MO.NMOVNRO AND MOI.NMOVITEM = MO.NMOVITEM " _
                & "         JOIN CTAIF CI           ON CI.CCTAIFCOD = MOI.CCTAIFCOD AND MOI.CPERSCOD = CI.CPERSCOD AND MOI.CIFTPO = CI.CIFTPO  " _
                & "         JOIN CONSTANTE C        ON C.nConsValor = SUBSTRING(CI.CCTAIFCOD,2,1) AND C.nCONSCOD LIKE '" & gCGTipoCtaIF & "' " _
                & "         JOIN INSTITUCIONFINANC I    ON I.CPERSCOD = CI.CPERSCOD AND I.CIFTPO = CI.CIFTPO " _
                & "         JOIN PERSONA P          ON P.CPERSCOD = I.CPERSCOD " _
                & "         JOIN CONSTANTE CTPOI        ON CTPOI.nConsValor = I.cIFTpo AND CTPOI.nCONSCOD LIKE '" & gCGTipoIF & "' " _
                & " WHERE M.cMovNro = '" & lsMovNro & "' AND MO.cObjetoCod ='" & Format(ObjEntidadesFinancieras, "00") & "'"
            
            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cpersNombre) & " " & Trim(rs!DescCtaIF)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cpersNombre) & " " & Trim(rs!DescCtaIF)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat("Cuenta(s) Origen", 20) & " :" + PrnSet("C+") + sOrigen + PrnSet("C-") + oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat("Cuenta(s) Destino", 20) & " :" + PrnSet("C+") + sDestino + PrnSet("C-") + oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
            sSQL = "    SELECT  DISTINCT M.CMOVNRO, " _
                & "             CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                     WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "             MOA.cAreaCod + MOA.CAGECOD AS cAreaCod, A.cAreaDescripcion + ' ' + ISNULL(AG.CAGEDESCRIPCION,'' )  as cAreaDescripcion , MO.COBJETOCOD, O.cObjetoDesc " _
                & "     FROM    MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "             JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "             JOIN MovObjAreaAgencia MOA         ON MOA.NMOVNRO = MO.NMOVNRO AND MOA.NMOVITEM = MO.NMOVITEM " _
                & "             JOIN AREAS A            ON A.cAreaCod = MOA.cAreaCod LEFT JOIN AGENCIAS AG ON AG.CAGECOD = MOA.CAGECOD " _
                & "             JOIN OBJETO O ON O.cObjetoCod  = MO.cObjetoCod " _
                & "     WHERE   M.cMovNro = '" & lsMovNro & "' AND Mo.cObjetoCod IN ('" & Format(ObjCMACAgenciaArea, "00") & "','" & Format(ObjCMACArea, "00") & "')"

            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                lsDesObjeto = Trim(rs!cObjetoDesc)
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!cAreaDescripcion)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!cAreaDescripcion)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Origen ", 20) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Destino ", 20) & " :" & sDestino & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
            sSQL = " SELECT    DISTINCT M.CMOVNRO, " _
                & "             CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                     WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "             MO.COBJETOCOD, O.cObjetoDesc " _
                & "   FROM      MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "             JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "             JOIN OBJETO O ON O.cObjetoCod  = MO.cObjetoCod " _
                & "   WHERE     M.cMovNro='" & lsMovNro & "' " _
                & "             AND MO.cObjetoCod NOT IN ('" & Format(ObjCMACAgencias, "00") & "','" & Format(ObjCMACAgenciaArea, "00") & "','" & Format(ObjDescomEfectivo, "00") & "','" & Format(ObjEntidadesFinancieras, "00") & "','" & Format(ObjBienesServicios, "00") & "','" & Format(ObjPersona, "00") & "')"
            
            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                lsDesObjeto = "OBJETOS"
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cObjetoCod) & "-" & Trim(rs!cObjetoDesc)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cObjetoCod) & "-" & Trim(rs!cObjetoDesc)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Origen ", 20) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Destino", 20) & " :" & sDestino & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
        End If ' del pscab1
    
       'Tipo de Cambio
        sSQL = "SELECT nMovTpoCambio, m.cOpeCod FROM MovTpoCambio MC JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE M.cMovNro = '" & lsMovNro & "'"
        Set rs = oConect.CargaRecordSet(sSQL)
        If Not rs.EOF Then
            If rs!cOpeCod = gContAjusteTipoCambio Then
            Else
               sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "      T.C.Mercado: " & Format(rs!nMovTpoCambio, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
            End If
            nLi = nLi + 1
        Else
            If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Then
                sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "     T.C.Compra : " & Format(lnCambioCompra, "##,###,##0.000") & "     T.C.Venta : " & Format(lnCambioVenta, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
        End If
       nLi = nLi + 1
       'Datos de Documentos
       sSQL = "SELECT b.cDocAbrev, a.cDocNro, a.dDocFecha FROM MovDoc a join mov m on m.nmovnro =a.nmovnro  , " & vsBaseComunes & "Documento b " _
            & "WHERE m.cMovNro = '" & lsMovNro & "' and b.nDocTpo = a.nDocTpo "
       
        Set rsDoc = oConect.CargaRecordSet(sSQL)
        If Not rsDoc.EOF Then
          sDoc = ""
          nLi = 1
          Do While Not rsDoc.EOF
             sDoc = sDoc & rsDoc!cDocAbrev & "_" & rsDoc!cDocNro & " " & rsDoc!dDocFecha & "  "
             If nLi Mod 4 = 0 Then
                sDoc = sDoc + oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
             End If
             rsDoc.MoveNext
          Loop
          sAsiento = sAsiento + ImpreFormat("Documentos", 20) & " :" + PrnSet("C+") + sDoc + PrnSet("C-") + oImpresora.gPrnSaltoLinea
          nLi = nLi + 1
        End If
        rsDoc.Close
       'Glosa
       sAsiento = sAsiento & ImpreGlosa(lsGlosa, pnColPage, ImpreFormat("Glosa       ", 20) & " :", , , nLi) + oImpresora.gPrnSaltoLinea
       nLi = nLi + 1
       lsMovBilletaje = ""
       lsMovBilletaje = ImprimeBilletaje(Mid(lsOpeCod, 3, 1), lsMovNro, pnColPage, , nLi)
       If lsMovBilletaje <> "" Then
            sAsiento = sAsiento + lsMovBilletaje
       End If
    'Cabecera 2 de Asiento
'    sCabecera = sCabecera + CON
'    sCabecera = sCabecera + "====================================================================================================================================" & oImpresora.gPrnSaltoLinea
'    sCabecera = sCabecera + "Itm C U E N T A   C O N T A B L E                                                                           DEBE            HABER" & oImpresora.gPrnSaltoLinea
'    sCabecera = sCabecera + "    Codigo              Descripcion" & oImpresora.gPrnSaltoLinea
'    sCabecera = sCabecera + "------------------------------------------------------------------------------------------------------------------------------------" & COFF & oImpresora.gPrnSaltoLinea
'    'MovCta para MN
'    sSQL = "SELECT mc.nMovItem, mc.cCtaContCod, ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),'') cCtaContDesc, mc.nMovImporte " _
'         & "FROM   MovCta mc JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE  M.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
'    Set Rs = oConect.CargaRecordSet(sSQL)
'    If Rs.EOF Then
'       'MsgBox "No existen datos de Movimiento ", vbInformation, "Error"
'       Exit Function
'    End If
'    nTot = 0
'    nTotH = 0
'    sAsiento = sAsiento + sCabecera
'    nLi = nLi + 4
'    sAsiento = sAsiento + BON & " EN MONEDA NACIONAL " + oImpresora.gPrnSaltoLinea
'    sAsiento = sAsiento + "--------------------" & BOFF + oImpresora.gPrnSaltoLinea
'    sAsiento = sAsiento & oImpresora.gPrnSaltoLinea
'    Do While Not Rs.EOF
'       DoEvents
'       If nLi > pnLinPage - 6 Then
'           nLi = 5
'           sAsiento = sAsiento + oImpresora.gPrnSaltoPagina & sCabecera
'       End If
'       sAsiento = sAsiento + CON & Rs!nMovItem & " " & Justifica(Rs!cCtaContCod, 22) & " " & Justifica(Rs!cCtaContDesc, 71) & " " _
'            & Right(Space(16) & IIf(Rs!nMovImporte > 0, Format(Rs!nMovImporte, "##,###,##0.00"), ""), 16) & " " _
'            & Right(Space(16) & IIf(Rs!nMovImporte < 0, Format(Rs!nMovImporte * -1, "##,###,##0.00"), ""), 16) _
'            & COFF & oImpresora.gPrnSaltoLinea
'       If Rs!nMovImporte > 0 Then
'          nTot = nTot + Val(Rs!nMovImporte)
'       Else
'          nTotH = nTotH + Val(Rs!nMovImporte) * -1
'       End If
'       nLi = nLi + 1
'       Rs.MoveNext
'    Loop
'    sAsiento = sAsiento + CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
'    sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF & oImpresora.gPrnSaltoLinea
'    nLi = nLi + 2
'
'    sSQL = "SELECT mc.nMovItem, mc.cCtaContCod, LTRIM(RTRIM(ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),''))) cCtaContDesc, me.nmovMEimporte " _
'         & "FROM   mov m join MovCta mc on mc.nmovnro=m.nmovnro JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'         & "WHERE  m.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
'    Set Rs = oConect.CargaRecordSet(sSQL)
'    sTexto = ""
'    nTot = 0
'    nTotH = 0
'    If Not Rs.EOF Then
'       sAsiento = sAsiento + BON & " EN MONEDA EXTRANJERA " & oImpresora.gPrnSaltoLinea
'       sAsiento = sAsiento + "----------------------" & BOFF & oImpresora.gPrnSaltoLinea
'       nLi = nLi + 2
'    End If
'    Do While Not Rs.EOF
'       DoEvents
'       If nLi > pnLinPage - 6 Then
'          nLi = 5
'          sTexto = sTexto + oImpresora.gPrnSaltoPagina
'          sTexto = sTexto + sCabecera
'       End If
'       sTexto = sTexto + CON & Rs!nMovItem & " " & Justifica(Rs!cCtaContCod, 22) & " " _
'              & Justifica(Rs!cCtaContDesc, 71) & " " _
'              & IIf(Rs!nMovMEImporte > 0, PrnVal(Rs!nMovMEImporte, 16, 2), Space(16)) & " " _
'              & IIf(Rs!nMovMEImporte < 0, PrnVal(Rs!nMovMEImporte * -1, 16, 2), Space(16)) _
'              & COFF & oImpresora.gPrnSaltoLinea
'       If Rs!nMovMEImporte > 0 Then
'          nTot = nTot + Val(Rs!nMovMEImporte)
'       Else
'          nTotH = nTotH + Val(Rs!nMovMEImporte) * -1
'       End If
'       nLi = nLi + 1
'       Rs.MoveNext
'    Loop
'    If nTot > 0 Or nTotH > 0 Then
'       sAsiento = sAsiento & sTexto
'       sAsiento = sAsiento + CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
'       sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF + oImpresora.gPrnSaltoLinea
'       nLi = nLi + 3
'    End If
'    If Rs.State = adStateOpen Then Rs.Close: Set Rs = Nothing
'    sAsiento = sAsiento + CON & "====================================================================================================================================" & COFF & oImpresora.gPrnSaltoLinea
    
    sAsiento = sAsiento + ImprePiePag(pnColPage, psPiePag, psOtraFirma)
    sAsiento = sAsiento + "   " + oImpresora.gPrnSaltoLinea

    Set oDGeneral = Nothing
    Set oConect = Nothing
    ImprimeAsientoNoContable = sAsiento
    
    Exit Function
ImprimeAsientoNoContableErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeAsientoNoContable Method")
End Function

Public Function ImprimeReciboEgresos(ByVal pnColPage As Integer, ByVal psMovNro As String, ByVal psGlosa As String, _
                                ByVal pdFecha As String, ByVal psEmpresaLogo As String, ByVal psOpecod As String, _
                                ByVal lbEsCajaChica As Boolean, ByVal psDescCajaChica As String, ByVal pnArendirFase As ARendirFases, _
                                ByVal psDocNroARendir As String, ByVal psDocNro As String, ByVal psFechaDoc As String, _
                                ByVal psPersCodProv As String, ByVal psPersNombreProv As String, _
                                ByVal psCargo As String, ByVal pnTotal As Currency) As String
    Dim sTexto As String
    Dim nAncho As Integer, N As Integer
    Dim lsTexto As String
    Dim BON As String
    Dim BOFF As String
    Dim CON As String
    Dim COFF As String
    On Error GoTo ImprimeReciboEgresosErr
    

    nAncho = pnColPage
    lsTexto = ""
    BON = PrnSet("B+")
    BOFF = PrnSet("B-")
    CON = PrnSet("C+")
    COFF = PrnSet("C-")
  
    lsTexto = lsTexto + ImpreCabAsiento(pnColPage, pdFecha, psEmpresaLogo, psOpecod, psMovNro, " R E C I B O   D E   E G R E S O  Nro. " & psDocNro, , , , 1) & oImpresora.gPrnSaltoLinea
'    lsTexto = lsTexto + ImpreCabAsiento(pnColPage, pdFecha, psEmpresaLogo, psOpeCod, psMovNro, " R E C I B O   D E   I N G R E S O S  Nro. " & psDocNro, , , , 1) & oImpresora.gPrnSaltoLinea
    If lbEsCajaChica Then
         lsTexto = lsTexto + "  CAJA CHICA   : " & BON & Mid(psDescCajaChica, 1, 60) & BOFF & oImpresora.gPrnSaltoLinea
    End If
    If pnArendirFase = ArendirRendicion Or pnArendirFase = ArendirSustentacion Then
        lsTexto = lsTexto + "  Rec. Arendir : " & BON & psDocNroARendir & BOFF & "  Fecha : " & BON & psFechaDoc & BOFF & oImpresora.gPrnSaltoLinea
    End If
    lsTexto = lsTexto + "  Persona      : " & BON & psPersCodProv & "  " & ImpreCarEsp(psPersNombreProv) & BOFF & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + "  Cargo        : " & psCargo & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + "  Importe      : " & BON & ConvNumLet(Trim(Str(pnTotal)), False, , Mid(psOpecod, 3, 1)) & BOFF & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + ImpreGlosa(psGlosa, pnColPage, "  Concepto     : ") & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + "" + oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + BON & Space(15) & "_________________________                ______________________" & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + Space(15) & CON & Centra(ImpreCarEsp(Mid(psPersNombreProv, 1, 40)), 40) & COFF & "                    Vo Bo JEFE AREA " & BOFF & oImpresora.gPrnSaltoLinea
  
    ImprimeReciboEgresos = lsTexto
    
    Exit Function
ImprimeReciboEgresosErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeReciboEgresos Method")
End Function

Public Function ImprimeRecibo(psMovNro As String, Optional pbDebe As Boolean = True, Optional psTit As String = "", Optional psCabe As String = "") As String
Dim sTexto As String, sTit As String
Dim N As Integer
Dim sSQL As String
Dim prs  As ADODB.Recordset
Dim oConect As DConecta
Set oConect = New DConecta
oConect.AbreConexion

sSQL = "SELECT IsNull(md.nDocTpo,'') cDocTpo, md.cDocNro, mop.cPersCod cProvCod, p.cPersNombre, ABS(ISNULL(ME.NMOVMEIMPORTE,mc.nMovImporte)) nMovImporte, m.cMovDesc, m.cOpecod, mch.cAreaCod + mch.cAgeCod cCHAreaCod, LTRIM(ISNULL(ag.cAgeDescripcion,'') + '  ' + a.cAreaDescripcion) cCHAreaDesc  " _
     & "FROM   Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro left join MOVME ME ON ME.NMOVNRO = MC.NMOVNRO AND ME.NMOVITEM = MC.NMOVITEM " _
     & "             JOIN MovObj mo ON mo.nMovNro = mc.nMovNro and mo.nMovItem = mc.nMovItem JOIN MovGasto mop ON mop.nMovNro = mo.nMovNro " _
     & "        LEFT JOIN MovDoc md ON md.nMovNro = m.nMovNro " _
     & "             JOIN Persona p ON p.cPersCod = mop.cPersCod LEFT JOIN MovCajaChica mch ON mch.nMovNro = m.nMovNro LEFT JOIN Areas a ON a.cAreaCod = mch.cAreaCod LEFT JOIN Agencias ag ON ag.cAgeCod = mch.cAgeCod " _
     & "WHERE  m.cMovNro = '" & psMovNro & "' and mc.nMovImporte " & IIf(pbDebe, ">", "<") & " 0 "
Set prs = oConect.CargaRecordSet(sSQL)
If prs.EOF Then
   RSClose prs
   Exit Function
End If
  sTexto = ""
  N = 1
  sTit = psTit
  Select Case prs!cDocTpo
      Case "49"
         Linea sTexto, ImpreCabAsiento(80, vsFecSis, vsEmpNom, prs!cOpeCod, psMovNro, " A U T O R I Z A C I O N   D E   E G R E S O  Nro. " & prs!cDocNro, , , , 1)
      Case TpoDocRecEgreso
         Linea sTexto, ImpreCabAsiento(80, vsFecSis, vsEmpNom, prs!cOpeCod, psMovNro, " R E C I B O   D E   E G R E S O  Nro. " & prs!cDocNro, , , , 1)
      Case Else
         Linea sTexto, ImpreCabAsiento(80, vsFecSis, vsEmpNom, prs!cOpeCod, psMovNro, sTit, , , , 1)
  End Select
  If sTit = "" And prs!cDocTpo = 0 Then
      sTit = " C O M P R O B A N T E "
      Linea sTexto, ImpreCabAsiento(80, vsFecSis, vsEmpNom, prs!cOpeCod, psMovNro, sTit, , , , 1)
  End If

  N = N + 4
  If psCabe <> "" Then
      Linea sTexto, psCabe
  End If
  If Not IsNull(prs!cCHAreaCod) Then
        Linea sTexto, "  CAJA CHICA   : " & prs!cCHAreaCod & " " & prs!cCHAreaDesc
  End If
  Linea sTexto, "  Persona      : " & BON & prs!cProvCod & "  " & ImpreCarEsp(prs!cpersNombre) & BOFF
  Linea sTexto, "  Cargo        : "
  Linea sTexto, "  Importe      : " & BON & ConvNumLet(prs!nMovImporte, False, , Mid(prs!cOpeCod, 3, 1)) & BOFF
  N = N + 3
  Linea sTexto, ImpreGlosa(prs!cMovdesc, gnColPage, "  Concepto     : ", 80, , N)
  Linea sTexto, "", 6
  N = N + 6
  Linea sTexto, BON & Space(15) & "_________________________                ______________________"
  N = N + 1
  Linea sTexto, Space(15) & CON & Centra(ImpreCarEsp(Mid(prs!cpersNombre, 1, 40)), 40) & COFF & "                    Vo Bo JEFE AREA " & BOFF
  N = N + 1
  Linea sTexto, " ", Int(gnLinPage / 2) - N
  RSClose prs
  
  ImprimeRecibo = sTexto
End Function

Public Function ImprimeReciboARendir(ByVal psMovNro As String, pnColPage As Integer, psEmpresaLogo As String, ByVal psEmpresa As String, psEmpresaRuc As String) As String
Dim sTexto As String
Dim lsCadPrint As String
Dim N As Integer
Dim BON As String
Dim BOFF As String
Dim oConect As DConecta
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsGlosa As String
Dim lsAgencia As String
Dim lsNomAge As String
Dim CON As String
Dim COFF As String

BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")
Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

sql = "Select   M.cMovNro,M.cMovDesc, MC.NMOVMONTO AS nMovSaldo , M.cOpeCod, " _
    & "         MD.nDocTpo, D.cDocDesc, MD.cDocNro, " _
    & "         ISNULL(A.cAreaCod,'') + ISNULL(AG.cAgeCod,'') AS cAreaCod, ISNULL(A.cAreaDescripcion,'') + ' ' + ISNULL(AG.cAgeDescripcion,'') AS cAreaDescripcion , MD.dDocFecha," _
    & "         ISNULL(AR.cAreaCod,'') + ISNULL(AGR.cAgeCod,'') AS cAreaCodPers, ISNULL(AR.cAreaDescripcion,'') + ' ' + ISNULL(AGR.cAgeDescripcion,'') AS cAreaDesPers , MD.dDocFecha," _
    & "         C.nConsValor  ,C.cConsDescripcion , P.cPersCod, P.cPersNombre, " _
    & "         ISNULL((Select cPersIDnro FROM PersID  WHERE cPersCod = P.cPersCod and cPersIDTpo =" & gPersIdDNI & "),'') AS DNI , " _
    & "         ISNULL((Select cPersIDnro FROM PersID  WHERE cPersCod = P.cPersCod and cPersIDTpo =" & gPersIdRUC & "),'') AS RUC ,  " _
    & "         MCH.cAreaCod + ' ' + MCH.cAgeCod as cAreaCH, ISNULL(MCH.nProcNro,'') AS nProcCh, A1.cAreaDescripcion + ' ' + ISNULL(AG1.cAgeDescripcion,'')   AS cDescCH " _
    & " From    Mov m JOIN MOVCONT MC ON MC.nMOVNRO=M.nMOVNRO   " _
    & "               Join MovARendir MR  ON MR.nMovNro = M.nMovNro " _
    & "               JOIN MovDoc MD      ON MD.nMovNro = MR.nMovNro " _
    & "               JOIN " & vsBaseComunes & "Documento D    ON D.nDocTpo = MD.nDocTpo " _
    & "          LEFT JOIN " & vsBaseComunes & "Areas A   ON A.cAreaCod = MR.cAreaCod " _
    & "          LEFT JOIN " & vsBaseComunes & "Agencias AG   ON AG.cAgeCod = MR.cAgeCod " _
    & "               JOIN " & vsBasePesonas & "Persona P      ON P.cPersCod = MR.cPersCod " _
    & "               JOIN RRHH RH      ON P.cPersCod = RH.cPersCod " _
    & "          LEFT JOIN " & vsBaseComunes & "Areas AR   ON AR.cAreaCod = RH.cAreaCodActual " _
    & "          LEFT JOIN " & vsBaseComunes & "Agencias AGR   ON AGR.cAgeCod = RH.cAgenciaActual    " _
    & "               JOIN " & vsBaseComunes & "Constante C    ON MR.cTpoARendir = C.nConsValor  AND C.nCONSCOD = '" & ConstanteCabecera.gArendirTipo & "' " _
    & "          LEFT JOIN MOVCAJACHICA MCH ON MCH.nMovNro= M.nMovNro " _
    & "          LEFT JOIN " & vsBaseComunes & "AREAS A1 ON A1.CAREACOD = MCH.cAreaCod " _
    & "          LEFT JOIN " & vsBaseComunes & "Agencias AG1   ON AG1.cAgeCod = MCH.cAgeCod" _
    & " Where M.cMovNro ='" & psMovNro & "' and M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "')"

Set rs = oConect.CargaRecordSet(sql)
If rs.EOF Then Exit Function

If rs!nConsValor = gArendirTipoCajaChica Then
    sTexto = "RECIBO DE A RENDIR CUENTA DE CAJA CHICA  NRO. " & rs!cDocNro
Else
    sTexto = "S O L I C I T U D   D E   F O N D O S   NRO. " & rs!cDocNro
End If
    lsCadPrint = ImpreCabAsiento(pnColPage, rs!dDocFecha, psEmpresaLogo, rs!cOpeCod, rs!cMovNro, sTexto, , , , 2)
    ' Definición de Cabecera 2
    If rs!nConsValor = gArendirTipoCajaChica Then
        lsCadPrint = lsCadPrint + "  CAJA CHICA   : " & BON & ImpreFormat(rs!cAreaCH & "-" & rs!nProcCh, 15) & ImpreCarEsp(rs!cDescCH) & BOFF & oImpresora.gPrnSaltoLinea
    Else
        lsCadPrint = lsCadPrint + "  A Rendir de  : " & BON & ImpreFormat(rs!cAreaCod, 15) & ImpreCarEsp(rs!cAreaDescripcion) & BOFF & oImpresora.gPrnSaltoLinea
    End If
    lsCadPrint = lsCadPrint + "  Persona      : " & BON & ImpreFormat(rs!cPerscod, 15) & ImpreCarEsp(PstaNombre(rs!cpersNombre)) & BOFF & oImpresora.gPrnSaltoLinea
    lsCadPrint = lsCadPrint + "  Area/Agencia : " & BON & ImpreFormat(rs!cAreaCodPers, 15) & ImpreCarEsp(rs!cAreaDesPers) & BOFF & oImpresora.gPrnSaltoLinea
    lsCadPrint = lsCadPrint + "  Cargo        : " & oImpresora.gPrnSaltoLinea
    lsCadPrint = lsCadPrint + "  Importe      : " & BON & ConvNumLet(rs!nMovSaldo, False, , Mid(rs!cOpeCod, 3, 1)) & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    If Mid(rs!cOpeCod, 3, 1) = gMonedaExtranjera Then
        'lsCadPrint = lsCadPrint + "                 Tipo de Cambio : " & txtTipCambio & oImpresora.gPrnSaltoLinea
    End If
    lsCadPrint = lsCadPrint + ImpreGlosa(Trim(rs!cMovdesc), pnColPage, "  Concepto     : ")
    lsCadPrint = lsCadPrint + oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadPrint = lsCadPrint + BON + oImpresora.gPrnSaltoLinea
    lsCadPrint = lsCadPrint + Space(20) & "___________________________" & oImpresora.gPrnSaltoLinea
    lsCadPrint = lsCadPrint + Space(20) & CON & Centra(Mid(PstaNombre(rs!cpersNombre), 1, 46), 46) & COFF & oImpresora.gPrnSaltoLinea
    lsCadPrint = lsCadPrint + Space(20) & Centra("  L.E. " & Trim(rs!DNI), 28) & BOFF & oImpresora.gPrnSaltoLinea
    lsCadPrint = lsCadPrint + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
    '***Modificado por ELRO el 20120524, según OYP-RFC005-2012 y OYP-RFC016-2012
    'lsGlosa = "De acuerdo a la Política existente, la rendición del presente recibo se deberá hacer en un plazo no mayor de 48 horas, bajo responsabilidad." & oImpresora.gPrnSaltoLinea _
    '       & "Los Comprobantes de gastos serán a nombre de la " & BON & Trim(psEmpresaLogo) & BOFF & ", RUC " & BON & psEmpresaRuc & BOFF & oImpresora.gPrnSaltoLinea & " "
    'lsCadPrint = lsCadPrint + ImpreGlosa(lsGlosa, pnColPage, " NOTA : ") + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
    'If rs!nConsValor <> gArendirTipoCajaChica Then
    '    lsCadPrint = lsCadPrint + Centra("____________________          _____________________", 80) + oImpresora.gPrnSaltoLinea
    '    lsCadPrint = lsCadPrint + Centra(" Vo Bo JEFE DE AREA              Vo Bo GERENCIA    ", 80) + oImpresora.gPrnSaltoLinea
    'Else
    '    lsCadPrint = lsCadPrint + Centra("____________________", 80) + oImpresora.gPrnSaltoLinea
    '   lsCadPrint = lsCadPrint + Centra(" Vo Bo JEFE DE AREA ", 80) + oImpresora.gPrnSaltoLinea
    'End If
    If rs!nConsValor = gArendirTipoCajaChica Then
        lsGlosa = "De acuerdo a la Política existente, la rendición del presente recibo se deberá hacer en un plazo no mayor de 48 horas, bajo responsabilidad." & oImpresora.gPrnSaltoLinea _
               & "Los Comprobantes de gastos serán a nombre de la " & BON & Trim(psEmpresaLogo) & BOFF & ", RUC " & BON & psEmpresaRuc & BOFF & oImpresora.gPrnSaltoLinea & " "
        lsCadPrint = lsCadPrint + ImpreGlosa(lsGlosa, pnColPage, " NOTA : ") + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
        lsCadPrint = lsCadPrint + Centra("____________________", 80) + oImpresora.gPrnSaltoLinea
        lsCadPrint = lsCadPrint + Centra(" Vo Bo JEFE DE AREA ", 80) + oImpresora.gPrnSaltoLinea
    End If
    '***Fin Modificado por ELRO*************************************************
    lsCadPrint = lsCadPrint + oImpresora.gPrnSaltoLinea
    ImprimeReciboARendir = lsCadPrint
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    Set oImp = Nothing

    Dim oCons As New NConstSistemas
    lsEmprLogo = oCons.LeeConstSistema(gConstSistNombreAbrevCMAC)
    gsNomCmac = oCons.LeeConstSistema(gConstSistCMACNombreCompleto)
    
    Set oCons = Nothing

Dim oIni As ClasIni
Set oIni = New ClasIni
vsBaseComunes = oIni.BaseComunes
vsBasePesonas = oIni.BasePersonas
sLPT = "LPT1"
BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")
Set oIni = Nothing
End Sub
Private Function ImprimeBilletaje(ByVal psMoneda As Moneda, ByVal psMovNro As String, ByVal pnColPage As Integer, Optional ByVal pbCondensado As Boolean = True, Optional ByRef nLin As Integer = 0) As String
Dim sSQL As String
Dim rs   As New ADODB.Recordset
Dim sTexto     As String
Dim sImpre As String
Dim BON As String
Dim BOFF As String
Dim COFF As String
Dim CON As String
Dim nMontoM    As Currency
Dim nMontoB As Currency
Dim nMontoTot  As Currency
Dim lsSimbolo As String
Dim otBilletajes() As tBilletajes
Dim lnFilaBill As Integer
Dim lnFilaMon As Integer
Dim lsMonedas As String
Dim lsBilletes As String
Dim I As Integer
Dim oEfec As Defectivo

BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")

On Error GoTo ErrBillete

Set oEfec = New Defectivo
ImprimeBilletaje = ""

sTexto = "": sImpre = ""
Set rs = oEfec.GetEfectivoMovImp(psMovNro, True)
If Not rs.EOF And Not rs.BOF Then
    If rs!Moneda = gMonedaExtranjera Then
        lsSimbolo = gcME
    Else
        '''lsSimbolo = gcMN 'marg ers044-2016
        lsSimbolo = gcPEN_SIMBOLO 'marg ers044-2016
        
    End If
    Linea sTexto, BON & PrnSet("N+") & "DESCOMPOSICION DE EFECTIVO : INGRESO" & BOFF & PrnSet("N-"), , nLin
    Linea sTexto, IIf(pbCondensado, CON, "") + String(pnColPage - 3, "="), , nLin
    nMontoM = 0: nMontoB = 0
    nMontoTot = 0
    Linea sTexto, BON & ImpreFormat(" BILLETES ", 22, 7) & ImpreFormat("MONTO", 15) _
         & ImpreFormat(" MONEDAS ", 14, 5) & ImpreFormat("MONTO", 10) + BOFF, , nLin
    Linea sTexto, String(pnColPage - 3, "-") + BOFF, , nLin
    
    ReDim otBilletajes(0)
    lnFilaBill = 0
    lnFilaMon = 0
    Do While Not rs.EOF
        If rs!billetes <> 0 Then
            lnFilaBill = lnFilaBill + 1
            If lnFilaBill > lnFilaMon Then
            ReDim Preserve otBilletajes(lnFilaBill)
            End If
            otBilletajes(lnFilaBill).lsBilletes = rs!billetes
            otBilletajes(lnFilaBill).lnMontoBil = rs!MontoBilletes
        End If
        If rs!Monedas <> 0 Then
            lnFilaMon = lnFilaMon + 1
            If lnFilaMon > lnFilaBill Then
            ReDim Preserve otBilletajes(lnFilaMon)
            End If
            otBilletajes(lnFilaMon).lsMonedas = rs!Monedas
            otBilletajes(lnFilaMon).lnMontoMon = rs!MontoMonedas
        End If
        nMontoM = nMontoM + rs!MontoMonedas
        nMontoB = nMontoB + rs!MontoBilletes
        nMontoTot = nMontoTot + rs!MontoMonedas + rs!MontoBilletes
       rs.MoveNext
    Loop
    For I = 1 To UBound(otBilletajes)
        If otBilletajes(I).lsBilletes <> 0 Then
            lsBilletes = ImpreFormat(lsSimbolo, 3, 3) & ImpreFormat(otBilletajes(I).lsBilletes, 7, 2, True) & ImpreFormat(otBilletajes(I).lnMontoBil, 18, 2, True)
        Else
            lsBilletes = ImpreFormat("", 34, 3)
        End If
        If otBilletajes(I).lsMonedas <> 0 Then
            lsMonedas = ImpreFormat(lsSimbolo, 6, 8) & ImpreFormat(otBilletajes(I).lsMonedas, 5, 2, True) & ImpreFormat(otBilletajes(I).lnMontoMon, 10, 2, True)
        Else
            lsMonedas = ImpreFormat("", 33, 3)
        End If
        Linea sTexto, lsBilletes + lsMonedas, , nLin
    Next
    Linea sTexto, String(pnColPage - 3, "-"), , nLin
    Linea sTexto, BON & ImpreFormat("SUBTOTAL :", 12, 3) & ImpreFormat(lsSimbolo, 3, 1) & ImpreFormat(nMontoB, 15, 2, True) _
         & ImpreFormat("SUBTOTAL", 10, 8) & ImpreFormat(lsSimbolo, 3, 1) & ImpreFormat(nMontoM, 10, 2, True) + BOFF, , nLin
    Linea sTexto, String(pnColPage - 3, "="), , nLin
    Linea sTexto, Space(45) & BON & "TOTAL :  " _
           & lsSimbolo & " " & PrnVal(nMontoTot, 14, 2) & BOFF, , nLin
    Linea sTexto, Space(45) & String(31, "=") + COFF, , nLin
End If
rs.Close
Set rs = Nothing

Set rs = oEfec.GetEfectivoMovImp(psMovNro, False)
If Not rs.EOF And Not rs.BOF Then
    If rs!Moneda = gMonedaExtranjera Then
        lsSimbolo = gcME
    Else
        '''lsSimbolo = gcMN 'marg ers044-2016
        lsSimbolo = gcPEN_SIMBOLO 'marg ers044-2016
    End If
    Linea sTexto, BON & PrnSet("N+") & "DESCOMPOSICION DE EFECTIVO : SALIDA" & BOFF & PrnSet("N-"), , nLin
    Linea sTexto, IIf(pbCondensado, CON, "") + String(pnColPage - 3, "="), , nLin
    nMontoM = 0: nMontoB = 0
    nMontoTot = 0
    Linea sTexto, BON & ImpreFormat(" BILLETES ", 22, 7) & ImpreFormat("MONTO", 15) _
         & ImpreFormat(" MONEDAS ", 14, 5) & ImpreFormat("MONTO", 10) + BOFF, , nLin
    Linea sTexto, String(pnColPage - 3, "-") + BOFF, , nLin
    ReDim otBilletajes(0)
    lnFilaBill = 0
    lnFilaMon = 0
    Do While Not rs.EOF
        If rs!billetes <> 0 Then
            lnFilaBill = lnFilaBill + 1
            If lnFilaBill > lnFilaMon Then
            ReDim Preserve otBilletajes(lnFilaBill)
            End If
            otBilletajes(lnFilaBill).lsBilletes = rs!billetes
            otBilletajes(lnFilaBill).lnMontoBil = rs!MontoBilletes
        End If
        If rs!Monedas <> 0 Then
            lnFilaMon = lnFilaMon + 1
            If lnFilaMon > lnFilaBill Then
            ReDim Preserve otBilletajes(lnFilaMon)
            End If
            otBilletajes(lnFilaMon).lsMonedas = rs!Monedas
            otBilletajes(lnFilaMon).lnMontoMon = rs!MontoMonedas
        End If
        nMontoM = nMontoM + rs!MontoMonedas
        nMontoB = nMontoB + rs!MontoBilletes
        nMontoTot = nMontoTot + rs!MontoMonedas + rs!MontoBilletes
       rs.MoveNext
    Loop
    For I = 1 To UBound(otBilletajes)
        If otBilletajes(I).lsBilletes <> 0 Then
            lsBilletes = ImpreFormat(lsSimbolo, 3, 3) & ImpreFormat(otBilletajes(I).lsBilletes, 7, 2, True) & ImpreFormat(otBilletajes(I).lnMontoBil, 18, 2, True)
        Else
            lsBilletes = ImpreFormat("", 34, 3)
        End If
        If otBilletajes(I).lsMonedas <> 0 Then
            lsMonedas = ImpreFormat(lsSimbolo, 6, 8) & ImpreFormat(otBilletajes(I).lsMonedas, 5, 2, True) & ImpreFormat(otBilletajes(I).lnMontoMon, 10, 2, True)
        Else
            lsMonedas = ImpreFormat("", 33, 3)
        End If
        Linea sTexto, lsBilletes + lsMonedas, , nLin
    Next
    Linea sTexto, String(pnColPage - 3, "-"), , nLin
    Linea sTexto, BON & ImpreFormat("SUBTOTAL :", 12, 3) & ImpreFormat(lsSimbolo, 3, 1) & ImpreFormat(nMontoB, 15, 2, True) _
         & ImpreFormat("SUBTOTAL", 10, 8) & ImpreFormat(lsSimbolo, 3, 1) & ImpreFormat(nMontoM, 10, 2, True) + BOFF, , nLin
    Linea sTexto, String(pnColPage - 3, "="), , nLin
    Linea sTexto, Space(45) & BON & "TOTAL :  " _
           & lsSimbolo & " " & PrnVal(nMontoTot, 14, 2) & BOFF, , nLin
    Linea sTexto, Space(45) & String(31, "=") + COFF, , nLin
End If
rs.Close
Set rs = Nothing
ImprimeBilletaje = sTexto
Exit Function
ErrBillete:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeBilletaje Method")
End Function
Public Function ImprimeOrdenPago(sTexto As String) As Boolean
Dim I As Integer
   ImprimeOrdenPago = True
   lbCancela = True
   frmImpresora.Show 1
   If lbCancela = False Then
      For I = 1 To lnNumCopias
         ImpreBegin False, 1
         Print #ArcSal, sTexto;
         Print #ArcSal, oImpresora.gPrnCondensadaOFF;    'Retorna al tipo de letra normal
         Close ArcSal
      Next I
   Else
      ImprimeOrdenPago = False
   End If
End Function
Public Function ImprimeReciboIngresoEgreso(ByVal psMovNro As String, pdFecha As Date, ByVal psGlosa As String, _
                                            ByVal psEmpresaLogo As String, ByVal psOpecod As String, _
                                            ByVal psPersCod As String, _
                                            ByVal lnMonto As Currency, ByVal pnColPage As Integer, _
                                            Optional pnTipoArendir As ArendirTipo, _
                                            Optional psNroRecViaticos As String, _
                                            Optional ByVal pbIngreso As Boolean = True, _
                                            Optional pbEsCajachica As Boolean, Optional ByVal psCHNro As String = "", _
                                            Optional ByVal psCHDesc As String = "", Optional pnArendirFase As ARendirFases = ArendirSolicitud, _
                                            Optional ByVal psDocNroARendir As String = "", Optional ByVal psFechaDoc As String = "", _
                                            Optional ByVal pbHabilitacionCH As Boolean = False) As String
Dim sTexto As String
Dim nAncho As Integer
Dim N As Integer
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim psAgeCod  As String, psAgeDesc As String
Dim psAreaCod  As String, psAreaDesc As String
Dim psPersNombre As String
    

Set oConect = New DConecta
Set rs = New ADODB.Recordset

If oConect.AbreConexion = False Then Exit Function

sql = " SELECT  P.cPersCod, P.cPersNombre , " _
    & "         ISNULL(A.cAreaCod,'') cAreaCod,  ISNULL(A.cAreaDescripcion,'') cAreaDescripcion, " _
    & "         ISNULL(AG.cAgeCod,'') cAgeCod, ISNULL(AG.cAgeDescripcion,'') cAgeDescripcion  " _
    & " FROM    PERSONA P " _
    & "         LEFT JOIN RRHH   RH ON RH.CPERSCOD = P.CPERSCOD " _
    & "         LEFT JOIN AREAS  A ON A.CAREACOD = RH.CAREACOD " _
    & "         LEFT JOIN AGENCIAS AG ON AG.CAGECOD = RH.cAgenciaActual " _
    & " WHERE   P.CPERSCOD ='" & psPersCod & "'"

Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    psAgeCod = rs!cAgeCod
    psAgeDesc = rs!cAgeDescripcion
    psAreaCod = rs!cAreaCod
    psAreaDesc = rs!cAreaDescripcion
    psPersNombre = PstaNombre(rs!cpersNombre)
Else
    Exit Function
End If
rs.Close
Set rs = Nothing

BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")

On Error GoTo ImprimeReciboIngresoEgresoErr
    nAncho = pnColPage
    sTexto = ""
    If pbHabilitacionCH = False Then
        If pbIngreso Then
            sTexto = sTexto + ImpreCabAsiento(pnColPage, pdFecha, psEmpresaLogo, psOpecod, psMovNro, " R E C I B O   D E   I N G R E S O ", , , , 1) & oImpresora.gPrnSaltoLinea
        Else
            sTexto = sTexto + ImpreCabAsiento(pnColPage, pdFecha, psEmpresaLogo, psOpecod, psMovNro, " R E C I B O   D E   E G R E S O ", , , , 1) & oImpresora.gPrnSaltoLinea
        End If
    Else
        sTexto = sTexto + ImpreCabAsiento(pnColPage, pdFecha, psEmpresaLogo, psOpecod, psMovNro, "H A B I L I T A C I O N    D E    C A J A    C H I C A", , , , 1) & oImpresora.gPrnSaltoLinea
    End If
    If pbEsCajachica Then
        sTexto = sTexto + "  CAJA CHICA   : " & BON & ImpreFormat(psCHNro, 12) & ImpreFormat(psCHDesc, 40) & BOFF & oImpresora.gPrnSaltoLinea
    End If
    If pnArendirFase = ArendirRendicion Or pnArendirFase = ArendirSustentacion Then
        sTexto = sTexto + "  Rec. Arendir : " & BON & ImpreFormat(psDocNroARendir, 14) & BOFF & "Fecha : " & BON & psFechaDoc & BOFF & oImpresora.gPrnSaltoLinea
    End If
    If pnTipoArendir = gArendirTipoViaticos Then
       sTexto = sTexto + "  Rec.Viaticos : " & BON & psNroRecViaticos & BOFF & oImpresora.gPrnSaltoLinea
    End If
    sTexto = sTexto + "  Agencia      : " & BON & ImpreFormat(psAgeCod, 14) & psAgeDesc & BOFF & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + "  Area         : " & BON & ImpreFormat(psAreaCod, 14) & psAreaDesc & BOFF & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + "  Persona      : " & BON & ImpreFormat(psPersCod, 14) & psPersNombre & BOFF & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + "  Importe      : " & BON & ConvNumLet(lnMonto, False, , Mid(psOpecod, 3, 1)) & BOFF & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + ImpreGlosa(psGlosa, pnColPage, "  Concepto     : ") & oImpresora.gPrnSaltoLinea
    'sTexto = sTexto + "" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    'sTexto = sTexto + BON & Space(20) & "_________________________" & oImpresora.gPrnSaltoLinea
    'Texto = sTexto + Space(20) & Centra("Vo Bo. CAJA ", 25) & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    If pbHabilitacionCH = False Then
        sTexto = sTexto + ImprePiePag(pnColPage, "1")
    Else
        sTexto = sTexto + ImprePiePag(pnColPage, "98", "Vo Bo GERENCIA")
    End If
    'End If
    ImprimeReciboIngresoEgreso = sTexto & oImpresora.gPrnSaltoPagina & sTexto
    Exit Function
    
ImprimeReciboIngresoEgresoErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeReciboIngresoEgreso Method")
End Function
Public Function ImprimePlanillaRendicion(ByVal psMovNroSol As String, ByVal pnColPage As Integer, ByVal pnLinPage As Integer, ByVal psNomCmac As String, _
                                         ByVal pdFecsis As Date, ByVal psCtaContArendir As String, ByVal psCtaContPendiente As String, ByVal psSimbolo As String, _
                                         ByVal pnMoneda As Moneda, ByVal pnTipoArendir As ArendirTipo, _
                                         Optional ByVal psOpecod As String = "") As String
                                         'Parametro psOpeCod agregado por ELRO el 20120425, según OYP-RFC005-2012
Dim N As Integer
Dim P As Integer
Dim nLin As Integer
Dim Cont As Integer
Dim lsTexto As String
Dim rs As ADODB.Recordset
Dim sql As String
Dim oNARendir As NARendir
Dim BON As String
Dim BOFF As String
Dim CON As String
Dim COFF As String
Dim lnSaldo As Currency
Dim lnSust As Currency

BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")


P = 0
nLin = 66
Cont = 0
Set oNARendir = New NARendir
Set rs = New ADODB.Recordset
'***Modificado por ELRO el 20120425, según OYP-RFC005-2012
'Set rs = oNARendir.GetDatosArendir(psMovNroSol, psCtaContArendir, psCtaContPendiente, pnMoneda, pnTipoArendir)
If psOpecod = gCGArendirViatSust2MN Or psOpecod = gCGArendirViatSust2ME Or psOpecod = gCGArendirViatRend2MN Or psOpecod = gCGArendirViatRend2ME Then
    Set rs = oNARendir.obtenerPlanillaRendicionSustentar(psMovNroSol, psCtaContArendir, pnMoneda)

ElseIf psOpecod = gCGArendirCtatSust2MN Or psOpecod = gCGArendirCtatSust2ME Or psOpecod = gCGArendirCtaRend2MN Or psOpecod = gCGArendirCtaRend2ME Then
    Set rs = oNARendir.obtenerPlanillaRendicionSustentar_2(psMovNroSol, psCtaContArendir, pnMoneda)

Else
    Set rs = oNARendir.GetDatosArendir(psMovNroSol, psCtaContArendir, psCtaContPendiente, pnMoneda, pnTipoArendir)
End If
'***Fin Modificado por ELRO*******************************

If Not rs.EOF And Not rs.BOF Then
    lnSaldo = rs!nMovSaldo
    lnSust = 0
    'Identificamos los Documentos Sustentatorios del Recibo de A rendir
    Do While Not rs.EOF
        lsTexto = lsTexto + CabeceraPlanillaRendicion(pnTipoArendir, " DECLARACION DE A RENDIR CUENTA ", nLin, P, pnLinPage, pnColPage, psNomCmac, pdFecsis, rs!cDocNro, IIf(IsNull(rs!cDocDescAtenc), "ATENCION", rs!cDocDescAtenc), _
                IIf(IsNull(rs!cDocNroAtenc), "EFECTIVO", rs!cDocNroAtenc), IIf(IsNull(rs!dFechaDocAtenc), Mid(rs!CMOVATENC, 7, 2) & "/" & Mid(rs!CMOVATENC, 5, 2) & "/" & Mid(rs!CMOVATENC, 1, 4), rs!dFechaDocAtenc), psSimbolo, rs!ImporteARendir, rs!cAgeCod, rs!cAgeDescripcion, _
                rs!cAreaCod, rs!cAreaDescripcion, rs!cPerscod, rs!cpersNombre, rs!cAreaCodArendir, rs!cAgeCodArendir, rs!cAreaArendir, rs!cAgeArendir)

        lsTexto = lsTexto + CON & "   " & Format(rs.Bookmark, "000") & "  " & rs!dDocFechaSust & "  " & _
                 Left(rs!cDocNroAbrevSust & Space(4), 4) & " " & Left(rs!cDocNroSust & Space(14), 14) & " " & _
                 Left(PstaNombre(rs!cPersCodSust) & Space(40), 40) & _
                 "  " & Left(rs!cDescSust & Space(35), 35) & " " & Right(Space(14) & Format(rs!nDocImporte, "#,##0.00"), 14) & COFF & oImpresora.gPrnSaltoLinea
                 
        nLin = nLin + 1
        lnSust = lnSust + rs!nDocImporte
        rs.MoveNext
    Loop
    lsTexto = lsTexto + CON
    lsTexto = lsTexto + " ------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + "                                                                              TOTAL SUSTENTADO                 :  " & psSimbolo & " " & PrnVal(CCur(Format(lnSust, "##,###0.00")), 14, 2) & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + "" & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + BON + "                                                                              SALDO POR RENDIR EN CAJA         :  " & psSimbolo & " " & PrnVal(CCur(Format(lnSaldo, "##,###0.00")), 14, 2) + BOFF + oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + " ====================================================================================================================================" & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + COFF + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + Centra("______________________                   ______________________ ", pnColPage) & oImpresora.gPrnSaltoLinea
    If pnTipoArendir = gArendirTipoViaticos Then
        lsTexto = lsTexto + Centra("   Vo Bo  USUARIO                          Vo Bo JEFE DE RRHH         ", pnColPage) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    Else
        lsTexto = lsTexto + Centra("   Vo Bo  USUARIO                          Vo Bo JEFE DE AREA         ", pnColPage) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    End If
End If
rs.Close
Set rs = Nothing
ImprimePlanillaRendicion = lsTexto
Set oNARendir = Nothing
End Function
'
'Public Function ImprimePlanillaRendicion(ByVal psMovNroSol As String, ByVal pnColPage As Integer, ByVal pnLinPage As Integer, ByVal psNomCmac As String, _
'                                         ByVal pdFecsis As Date, ByVal psCtaContArendir As String, ByVal psCtaContPendiente As String, ByVal psSimbolo As String, _
'                                         ByVal pnMoneda As Moneda, ByVal pnTipoArendir As ArendirTipo) As String
'Dim N As Integer
'Dim P As Integer
'Dim nLin As Integer
'Dim Cont As Integer
'Dim lsTexto As String
'Dim rs As ADODB.Recordset
'Dim sql As String
'Dim oNARendir As NARendir
'Dim BON As String
'Dim BOFF As String
'Dim CON As String
'Dim COFF As String
'Dim lnSaldo As Currency
'Dim lnSust As Currency
'
'BON = PrnSet("B+")
'BOFF = PrnSet("B-")
'CON = PrnSet("C+")
'COFF = PrnSet("C-")
'
'
'P = 0
'nLin = 66
'Cont = 0
'Set oNARendir = New NARendir
'Set rs = New ADODB.Recordset
'Set rs = oNARendir.GetDatosArendir(psMovNroSol, psCtaContArendir, psCtaContPendiente, pnMoneda, pnTipoArendir)
'If Not rs.EOF And Not rs.BOF Then
'    lnSaldo = rs!nMovSaldo
'    lnSust = 0
'    'Identificamos los Documentos Sustentatorios del Recibo de A rendir
'    Do While Not rs.EOF
'        lsTexto = lsTexto + CabeceraPlanillaRendicion(pnTipoArendir, " DECLARACION DE A RENDIR CUENTA ", nLin, P, pnLinPage, pnColPage, psNomCmac, pdFecsis, rs!cDocNro, IIf(IsNull(rs!cDocDescAtenc), "ATENCION", rs!cDocDescAtenc), _
'                IIf(IsNull(rs!cDocNroAtenc), "EFECTIVO", rs!cDocNroAtenc), IIf(IsNull(rs!dFechaDocAtenc), Mid(rs!CMOVATENC, 7, 2) & "/" & Mid(rs!CMOVATENC, 5, 2) & "/" & Mid(rs!CMOVATENC, 1, 4), rs!dFechaDocAtenc), psSimbolo, rs!ImporteARendir, rs!cAgeCod, rs!CAGEDESCRIPCION, _
'                rs!cAreaCod, rs!cAreaDescripcion, rs!cPerscod, rs!cpersNombre, rs!cAreaCodArendir, rs!cAgeCodArendir, rs!cAreaArendir, rs!cAgeArendir)
'
'        lsTexto = lsTexto + CON & "   " & Format(rs.Bookmark, "000") & "  " & rs!dDocFechaSust & "  " & _
'                 Left(rs!cDocNroAbrevSust & Space(4), 4) & " " & Left(rs!cDocNroSust & Space(14), 14) & " " & _
'                 Left(PstaNombre(rs!cPersCodSust) & Space(40), 40) & _
'                 "  " & Left(rs!cDescSust & Space(35), 35) & " " & Right(Space(14) & Format(rs!nDocImporte, "#,##0.00"), 14) & COFF & oImpresora.gPrnSaltoLinea
'
'        nLin = nLin + 1
'        lnSust = lnSust + rs!nDocImporte
'        rs.MoveNext
'    Loop
'    lsTexto = lsTexto + CON
'    lsTexto = lsTexto + " ------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
'    lsTexto = lsTexto + "                                                                              TOTAL SUSTENTADO                 :  " & psSimbolo & " " & PrnVal(CCur(Format(lnSust, "##,###0.00")), 14, 2) & oImpresora.gPrnSaltoLinea
'    lsTexto = lsTexto + "" & oImpresora.gPrnSaltoLinea
'    lsTexto = lsTexto + BON + "                                                                              SALDO POR RENDIR EN CAJA         :  " & psSimbolo & " " & PrnVal(CCur(Format(lnSaldo, "##,###0.00")), 14, 2) + BOFF + oImpresora.gPrnSaltoLinea
'    lsTexto = lsTexto + " ====================================================================================================================================" & oImpresora.gPrnSaltoLinea
'    lsTexto = lsTexto + COFF + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
'    lsTexto = lsTexto + Centra("______________________    ______________________   ______________________", pnColPage) & oImpresora.gPrnSaltoLinea
'    If pnTipoArendir = gArendirTipoViaticos Then
'        lsTexto = lsTexto + Centra("   Vo Bo  USUARIO           Vo Bo JEFE DE RRHH         Vo Bo GERENCIA    ", pnColPage) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'    Else
'        lsTexto = lsTexto + Centra("   Vo Bo  USUARIO           Vo Bo JEFE DE AREA         Vo Bo GERENCIA    ", pnColPage) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'    End If
'End If
'rs.Close
'Set rs = Nothing
'ImprimePlanillaRendicion = lsTexto
'Set oNARendir = Nothing
'End Function


Private Function CabeceraPlanillaRendicion(ByVal pnTipoArendir As ArendirTipo, ByVal psTitulo As String, ByRef nLin As Integer, ByRef P As Integer, ByVal pnLinPage As Integer, _
                                           ByVal pnColPage As Integer, ByVal psNomCmac As String, _
                                           ByVal pdFecsis As Date, ByVal psNroDocArendir As String, ByVal psDescDocAtend As String, _
                                           ByVal psNroDocAtend As String, ByVal psDocAtencFecha As String, _
                                           ByVal psSimbolo As String, ByVal ImporteARendir As Currency, _
                                           ByVal psAgeCod As String, ByVal psAgeDesc As String, _
                                           ByVal psAreaCod As String, ByVal psAreaDesc As String, _
                                           ByVal psPersCod As String, ByVal psPersNombre As String, _
                                           ByVal psAreaCodARendir As String, ByVal psAgeCodArendir As String, _
                                           ByVal psAreaDescArendir As String, ByVal psAgeDescArendir As String, _
                                           Optional pbCabRend As Boolean = True) As String
Dim lsTexto As String
Dim BON As String
Dim BOFF As String
Dim CON As String
Dim COFF As String

If nLin > pnLinPage - 4 Then
    BON = PrnSet("B+")
    BOFF = PrnSet("B-")
    CON = PrnSet("C+")
    COFF = PrnSet("C-")
   If P > 0 Then lsTexto = lsTexto & oImpresora.gPrnSaltoPagina
   P = P + 1
   lsTexto = lsTexto + ImpreFormat(psNomCmac, 60) & pdFecsis & " - " & Format(Time, "hh:mm:ss") + oImpresora.gPrnSaltoLinea
   lsTexto = lsTexto + " CONTABILIDAD    " & Space(55) & "Pag. " & Format(P, "000") + oImpresora.gPrnSaltoLinea
   lsTexto = lsTexto + BON & Centra(psTitulo, pnColPage) & BOFF + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
'   If pnTipoArendir = gArendirTipoCajaGeneral Then
        lsTexto = lsTexto + CON & "  Recibo de A rendir: " & BON & psNroDocArendir & BOFF + oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + "  A Rendir de.......: " & BON & psAreaCodARendir + psAgeCod & " " & psAreaDescArendir + " " + psAgeDescArendir & BOFF + oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + "  " & Mid(psDescDocAtend & Space(18), 1, 18) & ": " & BON & psNroDocAtend & BOFF & "       Fecha : " & BON & psDocAtencFecha & BOFF & _
            Space(15) & "Monto : " & BON & Right(Space(18) & psSimbolo & " " & Format(ImporteARendir, "#,#0.00"), 18) & BOFF + oImpresora.gPrnSaltoLinea
'   Else
'        lsTexto = lsTexto + CON & "  Recibo de A rendir: " & BON & psNroDocArendir & BOFF + Space(15) & "Monto : " & BON & Right(Space(18) & psSimbolo & " " & Format(ImporteARendir, "#,#0.00"), 18) & BOFF + oImpresora.gPrnSaltoLinea
'        lsTexto = lsTexto + "  A Rendir de.......: " & BON & psAreaCodARendir + psAgeCodArendir & " " & psAreaDescArendir + " " + psAgeDescArendir & BOFF + oImpresora.gPrnSaltoLinea
'   End If
   lsTexto = lsTexto + "  Usuario...........: " & BON & psPersCod & " " & PstaNombre(psPersNombre) & BOFF + oImpresora.gPrnSaltoLinea
   lsTexto = lsTexto + "  Area Funcional....: " & BON & psAreaCod & " " & psAreaDesc & BOFF + oImpresora.gPrnSaltoLinea
   If psAgeCod <> "" Then
        lsTexto = lsTexto + "  Agencia       ....: " & BON & psAgeCod & " " & psAgeDesc & BOFF + oImpresora.gPrnSaltoLinea
   End If
   lsTexto = lsTexto + oImpresora.gPrnSaltoLinea
   If pbCabRend Then
        lsTexto = lsTexto + " ====================================================================================================================================" & oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + "        DOCUMENTOS SUSTENTATORIOS          " & oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + "  ITEM  --------------------------------  PROVEEDOR                                 DETALLE                                  IMPORTE " & oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + "        Fecha      Tpo  Número                                                                                                       " & oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + " ------------------------------------------------------------------------------------------------------------------------------------" & COFF + oImpresora.gPrnSaltoLinea
    Else
        '''lsTexto = lsTexto + COFF + BON & "ASIENTO CONTABLE EN " & IIf(psSimbolo = "S/.", "M.N. :", "M.E. :") & BOFF & oImpresora.gPrnSaltoLinea 'marg ers044-2016
        lsTexto = lsTexto + COFF + BON & "ASIENTO CONTABLE EN " & IIf(psSimbolo = gcPEN_SIMBOLO, "M.N. :", "M.E. :") & BOFF & oImpresora.gPrnSaltoLinea 'marg ers044-2016
        lsTexto = lsTexto + String(pnColPage - 1, "=") & oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + "  " & BON & "Cuenta Contable" & Space(43) & "DEBE       HABER" & BOFF & oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + String(pnColPage - 1, "-") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    End If
    nLin = 15
End If
CabeceraPlanillaRendicion = lsTexto
End Function
Public Function ImprimeAsientoSustArendir(ByVal psMovNroAtenc As String, ByVal pnColPage As Integer, ByVal pnLinPage As Integer, ByVal psNomCmac As String, _
                                          ByVal pdFecsis As Date, ByVal psCtaContArendir As String, ByVal psSimbolo As String, ByVal pnTipoArendir As ArendirTipo, _
                                          Optional ByVal psOpecod As String = "")
                                          'Parametro psOpeCod agregado por ELRO el 20120425, según OYP-RFC005-2012
Dim lsTexto As String
Dim rs As ADODB.Recordset
Dim oArendir As NARendir

Dim oGen As DGeneral
Dim BON As String: Dim BOFF As String
Dim CON As String:   Dim COFF As String

Dim lnDebe As Currency: Dim lnHaber As Currency
Dim nLin As Integer
Dim P As Integer

Dim lsTexto8 As String, lsTexto8C As String, lnDebe8 As Currency, lnHaber8 As Currency '*** PEAC 20110325

Set oGen = New DGeneral

Set oArendir = New NARendir
Set rs = New ADODB.Recordset

BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")
nLin = 66
P = 0

lsTexto8 = "": lsTexto8C = ""
'***Modificado por ELRO el 20120425, según OYP-RFC005-2012
'Set rs = oArendir.GetAsientosSustARendir(psMovNroAtenc, psCtaContArendir, pnTipoArendir)
If psOpecod = gCGArendirViatSust2MN Or psOpecod = gCGArendirViatSust2ME Or psOpecod = gCGArendirViatRend2MN Or psOpecod = gCGArendirViatRend2ME Then
    Set rs = oArendir.obtenerPlanillaAsientoContablesSustentar(psMovNroAtenc, psCtaContArendir, Mid(psOpecod, 3, 1))

ElseIf psOpecod = gCGArendirCtatSust2MN Or psOpecod = gCGArendirCtatSust2ME Or psOpecod = gCGArendirCtaRend2MN Or psOpecod = gCGArendirCtaRend2ME Then
    Set rs = oArendir.obtenerPlanillaAsientoContablesSustentar_2(psMovNroAtenc, psCtaContArendir, Mid(psOpecod, 3, 1))

Else
    Set rs = oArendir.GetAsientosSustARendir(psMovNroAtenc, psCtaContArendir, pnTipoArendir)
End If
'***Fin Modificado por ELRO*******************************
If Not rs.EOF And Not rs.BOF Then
    Do While Not rs.EOF
        
        lsTexto = lsTexto + CabeceraPlanillaRendicion(pnTipoArendir, "SUSTENTACION DE A RENDIR CUENTA ", nLin, P, pnLinPage, pnColPage, psNomCmac, pdFecsis, rs!cDocNro, IIf(rs!cDocDescAtenc = "", "ATENCION", rs!cDocDescAtenc), _
               IIf(rs!cDocNroAtenc = "", "EFECTIVO", rs!cDocNroAtenc), IIf(IsNull(rs!dFechaDocAtenc), Mid(rs!CMOVATENC, 7, 2) & "/" & Mid(rs!CMOVATENC, 5, 2) & "/" & Mid(rs!CMOVATENC, 1, 4), rs!dFechaDocAtenc), psSimbolo, rs!ImporteARendir, rs!cAgeCod, rs!cAgeDescripcion, _
               rs!cAreaCod, rs!cAreaDescripcion, rs!cPerscod, rs!cpersNombre, rs!cAreaCodArendir, rs!cAgeCodArendir, rs!cAreaArendir, rs!cAgeArendir, False)

        If Left(rs!CtaContSust, 1) = "8" Then '*** PEAC 20110325
            '''lsTexto8 = lsTexto8 + CON + "   " & Justifica(rs!CtaContSust, 18) & "  " & Justifica(oGen.CuentaNombre(rs!CtaContSust), 66) & " " & gcMN & " " & PrnVal(rs!Debe, 14, 2) & " " & gcMN & " " & PrnVal(rs!Haber, 14, 2) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            lsTexto8 = lsTexto8 + CON + "   " & Justifica(rs!CtaContSust, 18) & "  " & Justifica(oGen.CuentaNombre(rs!CtaContSust), 66) & " " & gcPEN_SIMBOLO & " " & PrnVal(rs!Debe, 14, 2) & " " & gcPEN_SIMBOLO & " " & PrnVal(rs!Haber, 14, 2) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            lnDebe8 = lnDebe8 + rs!Debe
            lnHaber8 = lnHaber8 + rs!Haber
            rs.MoveNext
        Else
            '''lsTexto = lsTexto + CON + "   " & Justifica(rs!CtaContSust, 18) & "  " & Justifica(oGen.CuentaNombre(rs!CtaContSust), 66) & " " & gcMN & " " & PrnVal(rs!Debe, 14, 2) & " " & gcMN & " " & PrnVal(rs!Haber, 14, 2) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            lsTexto = lsTexto + CON + "   " & Justifica(rs!CtaContSust, 18) & "  " & Justifica(oGen.CuentaNombre(rs!CtaContSust), 66) & " " & gcPEN_SIMBOLO & " " & PrnVal(rs!Debe, 14, 2) & " " & gcPEN_SIMBOLO & " " & PrnVal(rs!Haber, 14, 2) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            lnDebe = lnDebe + rs!Debe
            lnHaber = lnHaber + rs!Haber
            nLin = nLin + 1
            rs.MoveNext
        End If
    Loop
    
    If lsTexto8 <> "" Then '*** PEAC 20110325
        lsTexto8C = oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        lsTexto8C = lsTexto8C + "A S I E N T O   D E   C U E N T A   D E   O R D E N" & oImpresora.gPrnSaltoLinea
        lsTexto8C = lsTexto8C + String(pnColPage - 1, "=") & oImpresora.gPrnSaltoLinea
        lsTexto8C = lsTexto8C + CON + "Cuenta Contable                                           DEBE       HABER" & oImpresora.gPrnSaltoLinea
        lsTexto8C = lsTexto8C + String(pnColPage - 1, "-") & oImpresora.gPrnSaltoLinea
        lsTexto8C = lsTexto8C + lsTexto8
        lsTexto8C = lsTexto8C + String(pnColPage - 1, "-") & oImpresora.gPrnSaltoLinea
        '''lsTexto8C = lsTexto8C + Space(75) & "TOTALES        " & gcMN & Right(Space(15) & Format(lnDebe8, "###,###,##0.00"), 15) & "  " & gcMN & Right(Space(15) & Format(lnHaber8, "###,###,##0.00"), 15) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
        lsTexto8C = lsTexto8C + Space(75) & "TOTALES        " & gcPEN_SIMBOLO & Right(Space(15) & Format(lnDebe8, "###,###,##0.00"), 15) & "  " & gcPEN_SIMBOLO & Right(Space(15) & Format(lnHaber8, "###,###,##0.00"), 15) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
        lsTexto8C = lsTexto8C + String(pnColPage - 1, "=") & oImpresora.gPrnSaltoLinea
    End If
    
    lsTexto = lsTexto + COFF + oImpresora.gPrnSaltoLinea + String(pnColPage - 1, "-") & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + CON
    '''lsTexto = lsTexto + Space(75) & "TOTALES        " & gcMN & Right(Space(15) & Format(lnDebe, "###,###,##0.00"), 15) & "  " & gcMN & Right(Space(15) & Format(lnHaber, "###,###,##0.00"), 15) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    lsTexto = lsTexto + Space(75) & "TOTALES        " & gcPEN_SIMBOLO & Right(Space(15) & Format(lnDebe, "###,###,##0.00"), 15) & "  " & gcPEN_SIMBOLO & Right(Space(15) & Format(lnHaber, "###,###,##0.00"), 15) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    lsTexto = lsTexto + COFF
    lsTexto = lsTexto + String(pnColPage - 1, "=") & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + ImprePiePag(pnColPage, "89", "VoBo  SOLICITANTE")
    lsTexto = lsTexto + lsTexto8C '*** PEAC 20110325
    
End If
rs.Close
Set rs = Nothing
Set oGen = Nothing
Set oArendir = Nothing
ImprimeAsientoSustArendir = lsTexto

End Function
Public Function ImprimeDetDocCtaCont(ByVal psMovNroSol As String, ByVal pnColPage As Integer, ByVal pnLinPage As Integer, _
                                  ByVal psCtaContArendir As String, ByVal psCtaContPendiente As String, ByVal psSimbolo As String, pnTipoArendir As ArendirTipo)

Dim lsTexto As String
Dim rs As ADODB.Recordset
Dim oArendir As NARendir

Dim oGen As DGeneral
Dim nLi    As Integer, P As Integer
Dim sCabe  As String
Dim BON As String: Dim BOFF As String
Dim CON As String:   Dim COFF As String
Dim lnDebe As Currency: Dim lnHaber As Currency
Dim lsCtaCont As String

Set oGen = New DGeneral
Set oArendir = New NARendir
Set rs = New ADODB.Recordset

BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")

lsTexto = ""

Set rs = oArendir.GetDetDocCtaCont(psMovNroSol, psCtaContArendir, psCtaContPendiente, pnTipoArendir)

If Not rs.EOF And Not rs.BOF Then
    sCabe = CON + String(125, "=") & oImpresora.gPrnSaltoLinea _
      & "  D O C U M E N T O               C O N C E P T O                                                DEBE           HABER" & oImpresora.gPrnSaltoLinea _
      & "TIPO  NUMERO        FECHA " & oImpresora.gPrnSaltoLinea _
      & String(125, "-") + COFF & oImpresora.gPrnSaltoLinea

    nLi = pnLinPage: P = 0
    lsTexto = PrnSet("MI", 4)
    If nLi > pnLinPage - 6 Then
        lsTexto = lsTexto + Cabecera("DETALLE DE DOCUMENTOS POR CUENTA CONTABLE", P, psSimbolo, pnColPage, sCabe, , gsNomCmac) & oImpresora.gPrnSaltoLinea
        nLi = 9
    End If
    Do While Not rs.EOF
        lsTexto = lsTexto + CON + "CUENTA CONTABLE: " & Justifica(rs!CtaContSust, 18) & " " & Justifica(oGen.CuentaNombre(rs!CtaContSust), 80) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        nLi = nLi + 2
        lsCtaCont = rs!CtaContSust
        lnDebe = 0: lnHaber = 0
        Do While lsCtaCont = rs!CtaContSust
            lnDebe = lnDebe + rs!Debe
            lnHaber = lnHaber + rs!Haber
            lsTexto = lsTexto + Justifica(rs!cDocAbrev, 3) & " " & Justifica(rs!cDocNro, 16) & " " & Justifica(rs!dDocFecha, 10) & "  " & Justifica(rs!cMovdesc, 60) & " " & PrnVal(rs!Debe, 14, 2) & " " & PrnVal(rs!Haber, 14, 2) & oImpresora.gPrnSaltoLinea
            nLi = nLi + 1
            If nLi > pnLinPage - 6 Then
                lsTexto = lsTexto + Cabecera("DETALLE DE DOCUMENTOS POR CUENTA CONTABLE", P, psSimbolo, pnColPage, sCabe, , gsNomCmac) & oImpresora.gPrnSaltoLinea
                nLi = 9
            End If
            rs.MoveNext
            If rs.EOF Then Exit Do
        Loop
        lsTexto = lsTexto + String(125, "-") & oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + Space(94) & PrnVal(lnDebe, 14, 2) & " " & PrnVal(lnHaber, 14, 2) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        nLi = nLi + 3
    Loop
End If
rs.Close
Set rs = Nothing
Set oGen = Nothing
Set oArendir = Nothing

ImprimeDetDocCtaCont = lsTexto
End Function

Public Function ImprimeNotaAbono(pdFecha As String, pnImporte As Currency, psGlosa As String, Optional lsCodCta As String, Optional psPersona As String = "", Optional sDocPagado As String = "") As String
Dim sTexto As String, sImpre As String
Dim nLen As Integer, N As Integer
Dim nLin As Integer
Dim sTextoAux As String

    sImpre = oImpresora.gPrnInicializa + oImpresora.gPrnTamLetra10CPI
    Linea sImpre, PrnSet("Esp", 6)
    Linea sImpre, PrnSet("B+") + Space(48), 7
    Linea sImpre, PrnSet("Esp", 1)
    Linea sImpre, Space(12) + IIf(Len(psPersona) > 45, CON, "") + psPersona + COFF, 2
    Linea sImpre, Space(52) & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Mid(pdFecha, 7, 4)
    Linea sImpre, PrnSet("EspN")
    Linea sImpre, PrnSet("Esp", 4)
    Linea sImpre, Space(48), 5
    Linea sImpre, Space(50) & lsCodCta
    Linea sImpre, PrnSet("EspN"), 3
    If Len(lsCodCta) = 12 Then
        '''Linea sImpre, Space(48) & IIf(Mid(lsCodCta, 6, 1) = Moneda.gMonedaNacional, gcMN, IIf(Mid(lsCodCta, 6, 1) = Moneda.gMonedaExtranjera, gcME, "")) & Right(Space(13) & Format(pnImporte, gsFormatoNumeroView), 13) 'MARG ERS044-2016
        Linea sImpre, Space(48) & IIf(Mid(lsCodCta, 6, 1) = Moneda.gMonedaNacional, gcPEN_SIMBOLO, IIf(Mid(lsCodCta, 6, 1) = Moneda.gMonedaExtranjera, gcME, "")) & Right(Space(13) & Format(pnImporte, gsFormatoNumeroView), 13) 'MARG ERS044-2016
    Else
        '''Linea sImpre, Space(48) & IIf(Mid(lsCodCta, 9, 1) = Moneda.gMonedaNacional, gcMN, IIf(Mid(lsCodCta, 9, 1) = Moneda.gMonedaExtranjera, gcME, "")) & Right(Space(13) & Format(pnImporte, gsFormatoNumeroView), 13) 'MARG ERS044-2016
        Linea sImpre, Space(48) & IIf(Mid(lsCodCta, 9, 1) = Moneda.gMonedaNacional, gcPEN_SIMBOLO, IIf(Mid(lsCodCta, 9, 1) = Moneda.gMonedaExtranjera, gcME, "")) & Right(Space(13) & Format(pnImporte, gsFormatoNumeroView), 13)
    End If
    sTextoAux = Trim(psGlosa) & " " & Trim(sDocPagado)
    nLen = 6
    'sTexto = Space(8) & JustificaTexto(gcGlosa & oImpresora.gPrnSaltoLinea & sDocPagado, 44 - nLen)
    sTexto = IIf(Len(sTextoAux) > 100, Space(4), Space(7)) & JustificaTexto(IIf(Len(sTextoAux) > 100, PrnSet("C+"), "") + sTextoAux, IIf(Len(sTextoAux) > 100, 61, 44) - nLen, IIf(Len(sTextoAux) > 100, 4, 0)) & PrnSet("C-")
    N = 0
    Do While True
       nLin = nLin + 1
       N = InStr(sTexto, oImpresora.gPrnSaltoLinea)
       If N > 0 Then
          sImpre = sImpre & Mid(sTexto, 1, N - 1) & oImpresora.gPrnSaltoLinea & Space(nLen)
          sTexto = Mid(sTexto, N + 1, Len(sTexto))
       End If
       If N = 0 Then
          sImpre = sImpre & sTexto & oImpresora.gPrnSaltoLinea
          Exit Do
       End If
    Loop
    Linea sImpre, "", 8 - nLin
    If Len(lsCodCta) = 12 Then
        '''Linea sImpre, Space(48) & IIf(Mid(lsCodCta, 6, 1) = "1", gcMN, IIf(Mid(lsCodCta, 6, 1) = "2", gcME, "")) & Right(Space(13) & Format(pnImporte, gsFormatoNumeroView), 13) 'MARG ERS044-2016
        Linea sImpre, Space(48) & IIf(Mid(lsCodCta, 6, 1) = "1", gcPEN_SIMBOLO, IIf(Mid(lsCodCta, 6, 1) = "2", gcME, "")) & Right(Space(13) & Format(pnImporte, gsFormatoNumeroView), 13) 'MARG ERS044-2016
    Else
        '''Linea sImpre, Space(48) & IIf(Mid(lsCodCta, 9, 1) = "1", gcMN, IIf(Mid(lsCodCta, 9, 1) = "2", gcME, "")) & Right(Space(13) & Format(pnImporte, gsFormatoNumeroView), 13) 'MARG ERS044-2016
        Linea sImpre, Space(48) & IIf(Mid(lsCodCta, 9, 1) = "1", gcPEN_SIMBOLO, IIf(Mid(lsCodCta, 9, 1) = "2", gcME, "")) & Right(Space(13) & Format(pnImporte, gsFormatoNumeroView), 13) 'MARG ERS044-2016
    End If
    ImprimeNotaAbono = ImpreCarEsp(sImpre) + PrnSet("B-")
End Function

Public Function ImprimeNotaCargoAbono(ByVal psNroNota As String, ByVal psGlosa As String, _
                                        ByVal pnImporte As Currency, _
                                        ByVal psPersNombre As String, ByVal psPersDireccion As String, _
                                        ByVal psPersUbiGeo As String, ByVal pdFecha As Date, ByVal psMoneda As Moneda, _
                                        ByVal psNroCtaAhorros As String, ByVal pnTpoDoc As TpoDoc, ByVal psAgeDesc As String, _
                                        ByVal psCodUser As String)
                                        
Dim lsTexto As String
Dim lsTitulo As String
Dim BON As String
Dim BOFF As String
Dim CON As String
Dim COFF As String
Dim lnColIzq As Long
Dim lncol As Long
Dim lsGlosa As String
Dim lsGlosaLin As String
Dim Char As String
Dim lnPos As Integer
Dim lnLineas As Integer
Dim lnImporteCargo As Currency
Dim lnImporteAbono As Currency
Dim lsTextoGlosa As String
Dim lsTextoTotal As String
Dim I As Integer
Dim lnLineasGlosa As Integer
lnColIzq = 13
lncol = 78

BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")

If pnTpoDoc = TpoDocNotaCargo Then
    lsTitulo = "NOTA DE CARGO N° " & psNroNota
    lnImporteCargo = pnImporte
    lnImporteAbono = 0
    lsTextoTotal = " TOTAL CARGADO "
Else
    lsTitulo = "NOTA DE ABONO N° " & psNroNota
    lnImporteAbono = pnImporte
    lnImporteCargo = 0
    lsTextoTotal = " TOTAL ABONADO "
End If
lsTexto = CON + BON + PrnSet("I+") + ImpreFormat(lsTitulo, lncol, 50) & ImpreFormat(lsTitulo, lncol, 10) + PrnSet("I-") + BOFF + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
lsTexto = lsTexto + ImpreFormat(Trim(psAgeDesc), lncol, 3) & ImpreFormat(Trim(psAgeDesc), lncol, lnColIzq) + oImpresora.gPrnSaltoLinea
lnLineas = lnLineas + 1
lsTexto = lsTexto + ImpreFormat(CentrarCadena(String(33, "=") + " Concepto " + String(35, "="), lncol), lncol, 3) + ImpreFormat(CentrarCadena(String(33, "=") + " Concepto " + String(35, "="), lncol), lncol, lnColIzq) & oImpresora.gPrnSaltoLinea
lnLineas = lnLineas + 1
'Impresion de la glosa
lsGlosa = JustificaTexto(psGlosa, lncol - 6)
lsGlosaLin = "": lnPos = 0: Char = ""
lnLineasGlosa = 0
For I = 1 To Len(lsGlosa)
    Char = Mid(lsGlosa, I, 1)
    lsGlosaLin = lsGlosaLin & Char
    If Char = oImpresora.gPrnSaltoLinea Or I = Len(lsGlosa) Then
        lsGlosaLin = Mid(lsGlosaLin, 1, IIf(I = Len(lsGlosa), Len(lsGlosaLin), Len(lsGlosaLin) - 1))
        lsTextoGlosa = lsTextoGlosa & ImpreFormat(Trim(lsGlosaLin), lncol, lnColIzq - 7) & ImpreFormat(Trim(lsGlosaLin), lncol, lnColIzq) & oImpresora.gPrnSaltoLinea
        lnLineas = lnLineas + 1
        lnLineasGlosa = lnLineasGlosa + 1
        If lnLineasGlosa = 3 Then
            Exit For
        End If
        lsGlosaLin = ""
    End If
Next
lsTexto = lsTexto + lsTextoGlosa
lsTexto = lsTexto + ImpreFormat(String(50, "_") & " CARGOS " & String(10, "_") & " ABONOS " & String(14, "_"), lncol, 3) + ImpreFormat(String(50, "_") & " CARGOS " & String(10, "_") & " ABONOS " & String(14, "_"), lncol, lnColIzq) & oImpresora.gPrnSaltoLinea
If lnImporteCargo > 0 Then
    lsTexto = lsTexto + ImpreFormat(ImpreFormat("IMPORTE :", 40) & ImpreFormat(lnImporteCargo, 14, 2, True) & ImpreFormat(0, 15, 2, True), lncol, 3) + ImpreFormat(ImpreFormat("IMPORTE :", 40) & ImpreFormat(lnImporteCargo, 14, 2, True) & ImpreFormat(0, 15, 2, True), lncol, lnColIzq) & oImpresora.gPrnSaltoLinea
Else
    lsTexto = lsTexto + ImpreFormat(ImpreFormat("IMPORTE :", 40) & ImpreFormat(0, 14, 2, True) & ImpreFormat(lnImporteAbono, 15, 2, True), lncol, 3) + ImpreFormat(ImpreFormat("IMPORTE :", 40) & ImpreFormat(0, 14, 2, True) & ImpreFormat(lnImporteAbono, 15, 2, True), lncol, lnColIzq) & oImpresora.gPrnSaltoLinea
End If
lsTexto = lsTexto + ImpreFormat(String(lncol, "_"), lncol, 3) + ImpreFormat(String(lncol, "_"), lncol, lnColIzq) & oImpresora.gPrnSaltoLinea
lsTexto = lsTexto + ImpreFormat("Sr(es) :", lncol, 3) + ImpreFormat("Sr(es) :", lncol, lnColIzq) & oImpresora.gPrnSaltoLinea
lsTexto = lsTexto + BON + ImpreFormat(PstaNombre(psPersNombre), lncol, lnColIzq + 5) + ImpreFormat(PstaNombre(psPersNombre), lncol, lnColIzq + 5) & oImpresora.gPrnSaltoLinea
lsTexto = lsTexto + ImpreFormat(psPersDireccion, lncol, lnColIzq + 5) + ImpreFormat(psPersDireccion, lncol, lnColIzq + 5) & oImpresora.gPrnSaltoLinea
lsTexto = lsTexto + ImpreFormat(psPersUbiGeo, lncol, lnColIzq + 5) + ImpreFormat(psPersUbiGeo, lncol, lnColIzq + 5) + BOFF + oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea

lsTexto = lsTexto + ImpreFormat(String(3, "-") & "FECHA" & String(13, "-") & " MONEDA " & String(5, "-") & " CUENTA " & String(15, "-") + BON + lsTextoTotal + BOFF & String(15, "-"), lncol + 4, 3) _
            + ImpreFormat(String(3, "-") & "FECHA" & String(13, "-") & " MONEDA " & String(5, "-") & " CUENTA " & String(15, "-") + BON + lsTextoTotal + BOFF & String(8, "-"), lncol + 4, lnColIzq) & oImpresora.gPrnSaltoLinea

'''lsTexto = lsTexto + ImpreFormat(String(3, " ") & Format(pdFecha & " " & Time, "mm/dd/yyyy hh:mm:ss") & String(3, " ") & IIf(psMoneda = gMonedaNacional, "SOLES", "DOLARES") & String(3, " ") & psNroCtaAhorros + BON + ImpreFormat(pnImporte, 15, 2, True) + BOFF + ImpreFormat(psCodUser, 5, 4), lncol + 4, 3)_
'''+ ImpreFormat(String(3, " ") & Format(pdFecha & " " & Time, "mm/dd/yyyy hh:mm:ss") & String(3, " ") & IIf(psMoneda = gMonedaNacional, "SOLES", "DOLARES") & String(3, " ") & psNroCtaAhorros + BON + ImpreFormat(pnImporte, 15, 2, True) + BOFF + ImpreFormat(psCodUser, 5, 4), lncol + 4, lnColIzq) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016

lsTexto = lsTexto + ImpreFormat(String(3, " ") & Format(pdFecha & " " & Time, "mm/dd/yyyy hh:mm:ss") & String(3, " ") & IIf(psMoneda = gMonedaNacional, StrConv(gcPEN_PLURAL, vbUpperCase), "DOLARES") & String(3, " ") & psNroCtaAhorros + BON + ImpreFormat(pnImporte, 15, 2, True) + BOFF + ImpreFormat(psCodUser, 5, 4), lncol + 4, 3) _
            + ImpreFormat(String(3, " ") & Format(pdFecha & " " & Time, "mm/dd/yyyy hh:mm:ss") & String(3, " ") & IIf(psMoneda = gMonedaNacional, StrConv(gcPEN_PLURAL, vbUpperCase), "DOLARES") & String(3, " ") & psNroCtaAhorros + BON + ImpreFormat(pnImporte, 15, 2, True) + BOFF + ImpreFormat(psCodUser, 5, 4), lncol + 4, lnColIzq) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            

lsTexto = lsTexto + ImpreFormat(String(lncol, "-"), lncol, 3) + ImpreFormat(String(lncol, "-"), lncol, lnColIzq) + COFF + oImpresora.gPrnSaltoLinea

'ImprimeNotaCargoAbono = lsTexto & oImpresora.gPrnSaltoPagina & lsTexto
ImprimeNotaCargoAbono = lsTexto
End Function



Public Function ImprimeMovCta(psCtaCod As String, sFechaIni As String, sFechaFin As String) As String
Dim prs As New ADODB.Recordset
Dim clsMov As New DMov
Dim sTexto As String
On Error GoTo ImprimeMovCtaErr

Set prs = clsMov.CargaMovCta(, "mc.cCtaContCod = '" & psCtaCod & "' and substring(cMovNro,1,8) BETWEEN '" & sFechaIni & "' and '" & sFechaFin & "'")
Set clsMov = Nothing

If prs.EOF Then
   MsgBox "No se registraron Movimientos de Cuenta Contable " & psCtaCod, vbInformation, "Aviso"
   prs.Close: Set prs = Nothing
   Exit Function
End If
Linea sTexto, CabeRepo(vsEmpNom, vsAgeNom, "Contabilidad", "", vsFecSis, "", "CAMBIO DE CUENTAS CONTABLES EN MOVIMIENTOS", _
            "    Movimiento              Item  Cuenta Contable           DEBE            HABER", _
            "  -------------------------------------------------------------------------------", _
            0, 79), 0
Do While Not prs.EOF
   Linea sTexto, "  " & prs!cMovNro & "  " & Format(prs!nMovItem, "000") & "  " & prs!cCtaContCod & "   " & PrnVal(prs!nDebe, 15, 2) & PrnVal(prs!nHaber, 15, 2)
   prs.MoveNext
Loop
prs.Close: Set prs = Nothing
ImprimeMovCta = sTexto
Exit Function
ImprimeMovCtaErr:
   Err.Raise Err.Number, "NContImprimir: ImprimeMovCta Módulo", Err.Description
End Function


Public Function ImprimirReciboViatico(ByVal rsDV As ADODB.Recordset, _
                                       ByVal rsConc As ADODB.Recordset, _
                                       ByVal rsAux As ADODB.Recordset, _
                                       ByVal pnColPage As Integer, ByVal pdFecsis As Date, ByVal psOpecod As String, _
                                       ByVal psEmpresaLogo As String, ByVal psSimbolo As String, _
                                       ByVal psNroDocViaticos As String, ByVal psNroDocRef As String, ByVal pdFechaDocRef As String, _
                                       ByVal psAreaCod As String, ByVal psAreaDesc As String, _
                                       ByVal psAgeCod As String, ByVal psAgeDesc As String, _
                                       ByVal psPersCod As String, ByVal psPersNombre As String, ByVal psDocPers As String, _
                                       ByVal psPersCargo As String, ByVal psCategoria As String, _
                                       ByVal pbSeleccion As Boolean, ByVal psMovNro As String, _
                                       ByVal psEmpresa As String, ByVal psEmpresaRuc As String) As String
Dim sTexto As String
Dim nAncho As Integer, N As Integer
Dim lsGlosa As String
Dim I As Integer
Dim lnImporte As Currency
Dim lsMovNro As String
Dim lnTotal As Currency
Dim lsTitulo As String
sTexto = ""
nAncho = pnColPage
With rsDV
    Do While Not .EOF
        If sTexto <> "" Then
            sTexto = sTexto + oImpresora.gPrnSaltoPagina
        End If
        If pbSeleccion = True Then
            Do While Not .EOF
                If !cMovNroDet <> psMovNro Then
                    .MoveNext
                    rsAux.MoveNext
                    If .EOF Then
                        ImprimirReciboViatico = sTexto
                        Exit Function
                    End If
                Else
                    Exit Do
                End If
            Loop
            If psMovNro = "" Then
                lsMovNro = " RECIBO DE VISUALIZACION "
            Else
                lsMovNro = psMovNro
            End If
        End If
        lsTitulo = "RECIBO DE VIATICOS A RENDIR CUENTA NRO. " & psNroDocViaticos
        sTexto = sTexto + PrnSet("MI", 5) & ImpreCabAsiento(pnColPage, pdFecsis, psEmpresaLogo, psOpecod, lsMovNro, lsTitulo) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Documento Refe.  : " & BON & Mid(psNroDocRef, 1, 50) & BOFF & Space(10) & "Fecha : " & BON & pdFechaDocRef & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Area             : " & BON & psAreaCod & "  " & ImpreCarEsp(psAreaDesc) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Agencia          : " & BON & psAgeCod & "  " & ImpreCarEsp(psAgeDesc) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Persona          : " & BON & psPersCod & "  " & ImpreCarEsp(psPersNombre) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Cargo            : " & BON & Mid(psPersCargo, 1, 50) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Categoria        : " & BON & Mid(psCategoria, 1, 50) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Destino          : " & BON & Justifica(Mid(!Lugar, 1, 50), 50) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Transporte Ida   : " & BON & Justifica(Mid(!Ida, 1, 50), 50) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Transporte Vuelta: " & BON & Justifica(Mid(!Vuelta, 1, 50), 50) & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Partida      : " & BON & !Partida & BOFF & Space(10) & "Nro. Días    : " & BON & PrnVal(!Dias, 5, 0) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Importe      : " & BON & ConvNumLet(!Importe, False) & BOFF & oImpresora.gPrnSaltoLinea
        lsGlosa = !Motivo
        sTexto = sTexto + ImpreGlosa(lsGlosa, pnColPage, "  Concepto     : ") & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "   ==================================================================" & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "     C O N C E P T O                                   IMPORTE (" & psSimbolo & ")" & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "   ------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
        lnTotal = 0
        Do While Not rsConc.EOF
            lnImporte = 0
            For I = 1 To rsAux.Fields.Count - 1
                If rsAux.Fields(I).Name = Trim(rsConc!Concepto) Then
                    lnImporte = rsAux.Fields(I).value
                    Exit For
                End If
            Next
            sTexto = sTexto + "      " & Justifica(rsConc!Descripcion, 45) & "  " & PrnVal(lnImporte, 14, 2) & oImpresora.gPrnSaltoLinea
            lnTotal = lnTotal + lnImporte
            rsConc.MoveNext
        Loop
        sTexto = sTexto + "   ------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + BON + "                                   T O T A L     " & psSimbolo & " " & PrnVal(lnTotal, 14, 2) + BOFF + oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "   ==================================================================  " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + Space(20) & Centra("_____________________________", 50) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + Space(20) & BON & Centra(Mid(psPersNombre, 1, 49), 50) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + Space(20) & Centra("  L.E. " & psDocPers, 50) & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        '***Comentado por ELRO el 20120524, según OYP-RFC005-2012
        'lsGlosa = "De acuerdo a la Política existente, la rendición del presente recibo se deberá realizar en el plazo establecido, bajo responsabilidad." & oImpresora.gPrnSaltoLinea _
        '       & "Los Comprobantes de gastos serán a nombre de " & BON & Trim(psEmpresa) & BOFF & ", RUC " & BON & Trim(psEmpresaRuc) & BOFF & oImpresora.gPrnSaltoLinea & " "
        'sTexto = sTexto + ImpreGlosa(lsGlosa, pnColPage, " NOTA : ") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        'sTexto = sTexto + Centra("____________________          _____________________", 80) & oImpresora.gPrnSaltoLinea
        'sTexto = sTexto + Centra("    Vo Bo RRHH                    Vo Bo GERENCIA    ", 80) & oImpresora.gPrnSaltoLinea
        'sTexto = sTexto + oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & PrnSet("MI", 0)
        '***Fin Comentado por ELRO*******************************
        If .EOF Then
            Exit Do
        End If
        rsAux.MoveNext
        .MoveNext
    Loop
    ImprimirReciboViatico = sTexto
End With
   
End Function
Public Function ImprimirReciboViaticoData(ByVal pnColPage As Integer, ByVal pdFecsis As Date, ByVal psOpecod As String, _
                                       ByVal psEmpresaLogo As String, ByVal psSimbolo As String, _
                                       ByVal psEmpresa As String, ByVal psEmpresaRuc As String, _
                                       ByVal psMovNroViat As String) As String
Dim sTexto As String
Dim nAncho As Integer, N As Integer
Dim lsGlosa As String
Dim I As Integer
Dim lnImporte As Currency
Dim lsMovNro As String
Dim lnTotal As Currency
Dim rs As ADODB.Recordset
Dim rsC As ADODB.Recordset
Dim oArendir As NARendir
Dim lsTitulo As String

Set oArendir = New NARendir
Set rs = New ADODB.Recordset
Set rsC = New ADODB.Recordset

sTexto = ""
nAncho = pnColPage
Set rs = oArendir.GetDatosTotalesViaticos(psMovNroViat)
sTexto = ""
With rs
    Do While Not .EOF
        lsMovNro = rs!cMovNro
        If sTexto <> "" Then
            sTexto = sTexto + oImpresora.gPrnSaltoPagina
        End If
        lsTitulo = "RECIBO DE VIATICOS A RENDIR CUENTA NRO. " & !DOCVIAT
        sTexto = sTexto + PrnSet("MI", 5) & ImpreCabAsiento(pnColPage, pdFecsis, psEmpresaLogo, !cOpeCod, lsMovNro, lsTitulo) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Documento Refe.  : " & BON & Mid(!NRODOCREF, 1, 50) & BOFF & Space(10) & "Fecha : " & BON & !dDocFecha & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  A Rendir de...   : " & BON & !cCodAreaArendir & !cCodAgeArendir & "  " & ImpreCarEsp(!cAreaArendir) & " " & ImpreCarEsp(!cAgeArendir) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Persona          : " & BON & !cPerscod & "  " & ImpreCarEsp(PstaNombre(!cpersNombre)) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Area             : " & BON & !cAreaCod & "  " & ImpreCarEsp(!cAreaDescripcion) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Agencia          : " & BON & !cAgeCod & "  " & ImpreCarEsp(!cAgeDescripcion) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Cargo            : " & BON & Mid(!cRHCargoDescripcion, 1, 50) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Categoria        : " & BON & Mid(!Categoria, 1, 50) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Destino          : " & BON & Justifica(Mid(!Destino, 1, 50), 50) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Transporte Ida   : " & BON & Justifica(Mid(!TransIda, 1, 50), 50) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Transporte Vuelta: " & BON & Justifica(Mid(!TrasnVuelta, 1, 50), 50) & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Partida      : " & BON & !dPartida & BOFF & Space(10) & "Nro. Días    : " & BON & PrnVal(!nMovViaticosDias, 5, 0) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Importe      : " & BON & ConvNumLet(!nMovMonto, False) & BOFF & oImpresora.gPrnSaltoLinea
        lsGlosa = !Motivo
        sTexto = sTexto + ImpreGlosa(lsGlosa, pnColPage, "  Concepto     : ") & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "   ==================================================================" & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "     C O N C E P T O                                   IMPORTE (" & psSimbolo & ")" & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "   ------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
        lnTotal = 0
        Set rsC = oArendir.GetDetalleCostosViaticos(rs!cMovNro)
        Do While Not rsC.EOF
            sTexto = sTexto + "      " & Justifica(rsC!Descripcion, 45) & "  " & PrnVal(rsC!Importe, 14, 2) & oImpresora.gPrnSaltoLinea
            lnTotal = lnTotal + rsC!Importe
            rsC.MoveNext
        Loop
        rsC.Close
        Set rsC = Nothing
        sTexto = sTexto + "   ------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + BON + "                                   T O T A L     " & psSimbolo & " " & PrnVal(lnTotal, 14, 2) + BOFF + oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "   ==================================================================  " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + Space(20) & Centra("_____________________________", 50) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + Space(20) & BON & Centra(Mid(PstaNombre(!cpersNombre), 1, 49), 50) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + Space(20) & Centra("  LE/DNI " & !DNI, 50) & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        '***Comentado por ELRO el 20120524, según OYP-RFC005-2012
        'lsGlosa = "De acuerdo a la Política existente, la rendición del presente recibo se deberá realizar en el plazo establecido, bajo responsabilidad." & oImpresora.gPrnSaltoLinea _
        '       & "Los Comprobantes de gastos serán a nombre de " & BON & Trim(psEmpresa) & BOFF & ", RUC " & BON & Trim(psEmpresaRuc) & BOFF & oImpresora.gPrnSaltoLinea & " "
        'sTexto = sTexto + ImpreGlosa(lsGlosa, pnColPage, " NOTA : ") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        'sTexto = sTexto + Centra("____________________          _____________________", 80) & oImpresora.gPrnSaltoLinea
        'sTexto = sTexto + Centra("    Vo Bo RRHH                    Vo Bo GERENCIA    ", 80) & oImpresora.gPrnSaltoLinea
        'sTexto = sTexto + oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & PrnSet("MI", 0)
         '***Fin  Comentado por ELRO******************************
        If .EOF Then
            Exit Do
        End If
        .MoveNext
    Loop
    ImprimirReciboViaticoData = sTexto
End With
   
End Function
Public Sub Inicio(psEmpNom As String, psAgeNom As String, psFecSis As String)
vsAgeNom = psAgeNom
vsEmpNom = psEmpNom
vsFecSis = psFecSis
End Sub
Public Function GetImprimeDocRendicionCH(ByVal pnLinPage As Integer, ByVal psNomCmact As String, ByVal pdFecsis As Date, _
                                        ByVal pnColPage As Integer, ByVal pbValido As Boolean, _
                                        ByVal psNroCH As String, ByVal pnNroProc As Integer, ByVal psCHDesc As String, _
                                        ByVal pdFechaApertReembolso As Date, ByVal pnMontoAsignado As Currency, _
                                        ByVal psPersCod As String, ByVal psPersNombre As String, _
                                        ByVal psCtaArendir As String, ByVal psCtaPendiente As String, _
                                        ByVal psCtaFondoFijo As String, ByVal psSimbolo As String, ByVal pnSaldo As Currency, _
                                        ByVal rsArendir As ADODB.Recordset, ByVal rsEgreDir As ADODB.Recordset) As String
Dim N As Integer
Dim P As Integer, nLin As Integer
Dim Cont As Integer
Dim sTexto As String
Dim nTotal   As Currency
Dim nTotalR  As Currency
Dim nSaldoR  As Currency
Dim sTexto1  As Currency
Dim nImporte As Currency
Dim nTotalE  As Currency
Dim nTotalI  As Currency
Dim rs As ADODB.Recordset
Dim oCh As NCajaChica
Dim lnTotalSust As Currency
Dim rsSustentos As New ADODB.Recordset '***Agregado por ELRO el 20120705, según OYP-RFC047-2012


Set rs = New ADODB.Recordset
Set oCh = New NCajaChica


P = 0
nLin = 66
Cont = 0
BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")

sTexto = PrnSet("MI", 5) & CabeceraRepo(psSimbolo, pnLinPage, psNomCmact, pdFecsis, pnColPage, _
                                         pbValido, psNroCH, pnNroProc, psCHDesc, _
                                         pdFechaApertReembolso, pnMontoAsignado, psPersCod, psPersNombre, nLin, P)
                                         
Linea sTexto, CON & BON & " EGRESOS CON A RENDIR : " & BOFF
Linea sTexto, " --------------------   " & COFF
nLin = nLin + 2
nTotalR = 0:   nTotalI = 0:   nTotalE = 0: nSaldoR = 0: lnTotalSust = 0
If Not rsArendir Is Nothing Then
    Do While Not rsArendir.EOF
        sTexto = sTexto & CabeceraRepo(psSimbolo, pnLinPage, psNomCmact, pdFecsis, pnColPage, _
                                       pbValido, psNroCH, pnNroProc, psCHDesc, _
                                       pdFechaApertReembolso, pnMontoAsignado, psPersCod, psPersNombre, nLin, P)
        nLin = nLin + 2
        Linea sTexto, CON & "  Recibo de A rendir: " & rsArendir![Nro Doc] & "   " & rsArendir!Fecha & "   " & _
              Space(15) & "Monto : " & Right(Space(18) & psSimbolo & " " & Format(rsArendir!Importe, "#,#0.00"), 18) & ImpreFormat(IIf(Val(rsArendir!NroCajaCh) = pnNroProc, "", " Caja Chica N° " & rsArendir!NroCajaCh), 20, 10) & COFF
        Linea sTexto, CON & "  Usuario...........: " & rsArendir!Solicitante & COFF
        If Val(rsArendir!NroCajaCh) = pnNroProc Then
            nTotalR = nTotalR + rsArendir!Importe
        End If
        nSaldoR = nSaldoR + rsArendir!Saldo
        'se buscan las sustentaciones
        Set rs = oCh.GetCHSustAtenArendir(psCtaArendir, psCtaPendiente, rsArendir!cMovNroAtenc)
        If Not rs.EOF Then
           Do While Not rs.EOF
              Cont = Cont + 1
              nLin = nLin + 1
              sTexto = sTexto & CabeceraRepo(psSimbolo, pnLinPage, psNomCmact, pdFecsis, pnColPage, _
                                             pbValido, psNroCH, pnNroProc, psCHDesc, _
                                             pdFechaApertReembolso, pnMontoAsignado, psPersCod, psPersNombre, nLin, P)
                                             
              Linea sTexto, CON + PrnSet("I+") & ImpreFormat(Format(Cont, "000"), 3, 5) & ImpreFormat(rs!FECHADOC, 10) & _
                       ImpreFormat(rs!cDocAbrev, 3) & ImpreFormat(rs!cDocNro, 15) & ImpreFormat(PstaNombre(rs!cpersNombre), 35) & _
                       ImpreFormat(rs!Glosa, 32) & ImpreFormat(rs!Importe, 12, 2, True) & PrnSet("I-") + COFF
              lnTotalSust = lnTotalSust + Abs(rs!Importe)
              rs.MoveNext
           Loop
        End If
        rs.Close
        Set rs = Nothing
        If Val(rsArendir!Saldo) = 0 Then
            '***Agregado por ELRO el 20120628, según OYP-RFC047-2012
            'Set rs = oCh.GetDatosRendArendirCH(rsArendir!cMovNroAtenc, psCtaFondoFijo)
            '***Modificado por ELRO el 20130222, según SATI INC1301300007
            'Set rs = oCh.GetDatosRendArendirCH(rsArendir!cMovNroAtenc, "111701")
            Set rs = oCh.GetDatosRendArendirCH(rsArendir!cMovNroAtenc, psCtaFondoFijo)
            '***Fin Modificado por ELRO el 20130222**********************
            '***Fin Agregado por ELRO*******************************
            If Not rs.EOF Then
               Cont = Cont + 1
               nLin = nLin + 1
               CabeceraRepo psSimbolo, pnLinPage, psNomCmact, pdFecsis, pnColPage, _
                             pbValido, psNroCH, pnNroProc, psCHDesc, _
                             pdFechaApertReembolso, pnMontoAsignado, psPersCod, psPersNombre, nLin, P
                             
               Linea sTexto, CON + BON + ImpreFormat(Format(Cont, "000"), 3, 5) & ImpreFormat(rs!FechaRend, 10) & _
                            ImpreFormat(IIf(rs!nMovImporte > 0, "R/I ", "R/E"), 3) & ImpreFormat(rs!cMovNro, 52) & ImpreFormat(IIf(rs!nMovImporte > 0, "VUELTO INGRESADO A CAJA", "REEMBOLSO A USUARIO"), 30) & " " & _
                            ImpreFormat(Abs(rs!nMovImporte), 13, 2, True) + BOFF + COFF
                            
               If rs!nMovImporte > 0 Then
                     nTotalI = nTotalI + Abs(rs!nMovImporte)
               Else
                     nTotalE = nTotalE + Abs(rs!nMovImporte)
               End If
            End If
            rs.Close
            Set rs = Nothing
        Else
           nImporte = Format(rsArendir!Saldo, "#,#0.00")
           Linea sTexto, CON + PrnSet("I+") + ImpreFormat(Format(Cont, "000"), 74, 5) & _
                ImpreFormat(IIf(nImporte < 0, "REEMBOLSO PENDIENTE A USUARIO", "RENDICION PENDIENTE"), 30) & " " & _
                ImpreFormat(nImporte, 13, 2, True) + PrnSet("I-") + COFF
           nLin = nLin + 1
        End If
        rsArendir.MoveNext
    Loop

    Linea sTexto, CON & " ------------------------------------------------------------------------------------------------------------------------------------" & COFF
    Linea sTexto, CON & "   " & Space(60) & "TOTAL EGRESOS SUSTENTADOS    " & psSimbolo & " " & PrnVal(lnTotalSust, 16, 2)    ' nTotalR - nTotalI + nTotalE
    Linea sTexto, CON & "   " & Space(60) & "TOTAL INGRESOS POR RENDICION " & psSimbolo & " " & PrnVal(nTotalI, 16, 2)
    Linea sTexto, CON & "   " & Space(60) & "TOTAL EGRESOS  POR RENDICION " & psSimbolo & " " & PrnVal(nTotalE, 16, 2)
    Linea sTexto, CON & "   " & Space(60) & "---------------------------------------------------------------------"
    Linea sTexto, CON & "   " & Space(80) & "TOTAL EGRESOS POR A RENDIR   " & psSimbolo & " " & PrnVal(nTotalR, 16, 2)
    Linea sTexto, CON & "   " & Space(80) & "TOTAL A RENDIR PENDIENTES    " & psSimbolo & " " & PrnVal(nSaldoR, 16, 2)
    nLin = nLin + 9
End If

'***Agregado por ELRO el 2012075, según OYP-RFC047-2012
If lnTotalSust > 0 Then
    Linea sTexto, BON & " DOCUMENTOS SUSTENTATORIOS : " & BOFF
    Linea sTexto, " ----------------   " & COFF
    Cont = 0: nTotal = 0
    
    rsArendir.MoveFirst
    Do While Not rsArendir.EOF
        Set rsSustentos = oCh.GetCHSustAtenArendir(psCtaArendir, psCtaPendiente, rsArendir!cMovNroAtenc)
        If Not rsSustentos Is Nothing Then
            Do While Not rsSustentos.EOF
                  sTexto = sTexto & CabeceraRepo(psSimbolo, pnLinPage, psNomCmact, pdFecsis, pnColPage, _
                                                pbValido, psNroCH, pnNroProc, psCHDesc, _
                                                pdFechaApertReembolso, pnMontoAsignado, psPersCod, psPersNombre, nLin, P)
                  nLin = nLin + 1
                  Cont = Cont + 1
                  Linea sTexto, CON & "   " & Format(Cont, "000") & " " & Justifica(rsSustentos!cDocAbrev, 4) & " " & Justifica(rsSustentos!FECHADOC, 10) & " " & _
                                     Justifica(rsSustentos![cDocNro], 16) & " " & Justifica(PstaNombre(rsSustentos!cpersNombre), 35) & " " & Justifica(rsSustentos!Glosa, 41) & " " & PrnVal(rsSustentos!Importe, 14, 2) & COFF
                  nTotal = nTotal + rsSustentos!Importe
               rsSustentos.MoveNext
            Loop
        End If
    rsArendir.MoveNext
    Loop
    
    Linea sTexto, CON & " ------------------------------------------------------------------------------------------------------------------------------------" & COFF
    Linea sTexto, CON & "   " & Space(80) & "TOTAL DOCUMENTOS SUSTENTATORIOS " & psSimbolo & " " & PrnVal(nTotal, 16, 2), 2
End If
'rsArendir.Close'***Comentado por ELRO EL 20120813
Set rsArendir = Nothing
'***Fin Agregado por ELRO******************************

If Not rsEgreDir Is Nothing Then
    Linea sTexto, BON & " EGRESOS DIRECTOS : " & BOFF
    Linea sTexto, " ----------------   " & COFF
    Cont = 0: nTotal = 0
    Do While Not rsEgreDir.EOF
          sTexto = sTexto & CabeceraRepo(psSimbolo, pnLinPage, psNomCmact, pdFecsis, pnColPage, _
                                        pbValido, psNroCH, pnNroProc, psCHDesc, _
                                        pdFechaApertReembolso, pnMontoAsignado, psPersCod, psPersNombre, nLin, P)
          nLin = nLin + 1
          Cont = Cont + 1
          Linea sTexto, CON & "   " & Format(Cont, "000") & " " & Justifica(rsEgreDir!Tipo, 4) & " " & Justifica(rsEgreDir!Fecha, 10) & " " & _
                             Justifica(rsEgreDir![Nro Doc], 16) & " " & Justifica(PstaNombre(rsEgreDir!Solicitante), 35) & " " & Justifica(rsEgreDir!Concepto, 41) & " " & PrnVal(rsEgreDir!Importe, 14, 2) & COFF
          nTotal = nTotal + rsEgreDir!Importe
       rsEgreDir.MoveNext
    Loop
    Linea sTexto, CON & " ------------------------------------------------------------------------------------------------------------------------------------" & COFF
    Linea sTexto, CON & "   " & Space(80) & "TOTAL EGRESOS DIRECTOS       " & psSimbolo & " " & PrnVal(nTotal, 16, 2), 2
End If
Linea sTexto, " ------------------------------------------------------------------------------------------------------------------------------------"
Linea sTexto, "                                                                                         SALDO DE CAJA CHICA    " & BON & psSimbolo & " " & PrnVal(Format(pnSaldo, "#,#0.00"), 16, 2) & BOFF
Linea sTexto, " ===================================================================================================================================="
Linea sTexto, COFF, 6
Linea sTexto, Centra("_______________________          _______________________", pnColPage)
Linea sTexto, Centra("  Vo Bo Contabilidad                  Vo Bo Gerencia    ", pnColPage)
GetImprimeDocRendicionCH = sTexto

End Function
Private Function CabeceraRepo(ByVal psSimbolo As String, ByVal pnLinPage As Integer, ByVal psNomCmact As String, ByVal pdFecsis As Date, _
                              ByVal pnColPage As Integer, ByVal pbValido As Boolean, _
                              ByVal psNroCH As String, ByVal pnNroProc As Integer, ByVal psCHDesc As String, _
                              ByVal pdFechaApertReembolso As Date, ByVal pnMontoAsignado As Currency, _
                              ByVal psPersCod As String, ByVal psPersNombre As String, _
                              ByRef nLin As Integer, ByRef P As Integer) As String
Dim sCabe As String
sCabe = ""
If nLin > pnLinPage - 6 Then
   If P > 0 Then sCabe = sCabe & oImpresora.gPrnSaltoPagina
   P = P + 1
   Linea sCabe, CON + ImpreFormat(psNomCmact, 60, 0) & Space(42) & pdFecsis & " - " & Format(Time, "hh:mm:ss")
   Linea sCabe, ImpreFormat("CONTABILIDAD", 60, 0) & Space(48) & "Pag. " & Format(P, "000") & COFF
   Linea sCabe, BON & Centra(" DOCUMENTOS SUSTENTATORIOS DE CAJA CHICA ", pnColPage) & BOFF
   Linea sCabe, CON
   If pbValido = False Then
      Linea sCabe, BON & Centra("*** REPORTE NO VALIDO PARA RENDICION EN CONTABILIDAD ***", 80) & BOFF, 2
   End If
   If P = 1 Then
      Linea sCabe, " Area Responsable    : " & BON & ImpreFormat("[" & psNroCH & "] " & psCHDesc, 40, 0) & ImpreFormat("Caja Chica N° [" & pnNroProc & "]", 28) & BOFF & _
                    " Apertura/Reembolso : " & BON & pdFechaApertReembolso & BOFF
      Linea sCabe, " Persona Responsable : " & BON & Left("[" & psPersCod & "] " & psPersNombre & Space(70), 70) & BOFF & _
               "     Monto Asignado : " & psSimbolo & " " & BON & Format(pnMontoAsignado, "#,#0.00") & BOFF
      nLin = 10
   Else
      nLin = 8
   End If
   Linea sCabe, " ===================================================================================================================================="
   Linea sCabe, "   Item  Fecha de   Comprobante          Proveedor                                 Detalle                                    Monto  "
   Linea sCabe, "         Emisión    Tpo  Número                                                                         "
   Linea sCabe, " ------------------------------------------------------------------------------------------------------------------------------------" & COFF
End If
CabeceraRepo = sCabe
End Function
Public Function ImprimeAsientoRendicionCH(ByVal psAreaCh As String, ByVal psAgeCh As String, _
                              ByVal pnProcCH As Integer, ByVal psSimbolo As String, _
                              ByVal pnColPage As Integer, ByVal pnLinPage As Integer, ByVal pdFecsis As Date, _
                              ByVal psEmpresaLogo As String, ByVal psOpecod As String, _
                              ByVal psMovNro As String, ByVal pbValido As Boolean, ByVal psCtaFondoFijo As String) As String
Dim sTexto   As String
Dim sAsiento As String
Dim sMovs    As String
Dim sMov     As String
Dim sMov1    As String
Dim sUltMovRend As String
Dim sFecAsiento As String
Dim rs As ADODB.Recordset
Dim oCh As NCajaChica
Dim oGen As DGeneral
Dim lsImp As String
Dim lnTotalDebe As Currency
Dim lnTotalHaber As Currency
Dim nLi    As Integer, P As Integer
Dim sCabe  As String

Dim sTexto8 As String, lnTotalDebe8 As Double, lnTotalHaber8 As Double '*** PEAC 20110425
Dim sTexto8c As String
Set oGen = New DGeneral
Set rs = New ADODB.Recordset
Set oCh = New NCajaChica

Linea sAsiento, ImpreCabAsiento(pnColPage, pdFecsis, psEmpresaLogo, psOpecod, psMovNro, "", True, Format(pdFecsis, gsFormatoFechaView), False, 2)
If pbValido = False Then
   Linea sAsiento, BON & Centra("*** REPORTE NO VALIDO PARA RENDICION EN CONTABILIDAD ***", 80) & BOFF, 2
End If

Set rs = oCh.GetAsientoCH(psAreaCh, psAgeCh, pnProcCH)
lnTotalDebe = 0: lnTotalDebe8 = 0: lnTotalHaber8 = 0
lnTotalHaber = 0
sTexto = "": sTexto8 = "": sTexto8c = ""
Do While Not rs.EOF

    If Left(rs!cCtaContCod, 1) = "8" Then '*** PEAC 20110425
        Linea sTexto8, ImpreFormat(rs!cCtaContCod, 19, 5) & "  " & ImpreFormat(rs!cCtaContDesc, 60) & ImpreFormat(psSimbolo, 3) & ImpreFormat(rs!Debe, 12, 2, True) & ImpreFormat(psSimbolo, 3) & ImpreFormat(rs!Haber, 12, 2, True)
        lnTotalDebe8 = lnTotalDebe8 + rs!Debe
        lnTotalHaber8 = lnTotalHaber8 + rs!Haber
        rs.MoveNext
    Else
        Linea sTexto, ImpreFormat(rs!cCtaContCod, 19, 5) & "  " & ImpreFormat(rs!cCtaContDesc, 60) & ImpreFormat(psSimbolo, 3) & ImpreFormat(rs!Debe, 12, 2, True) & ImpreFormat(psSimbolo, 3) & ImpreFormat(rs!Haber, 12, 2, True)
        lnTotalDebe = lnTotalDebe + rs!Debe
        lnTotalHaber = lnTotalHaber + rs!Haber
        rs.MoveNext
    End If
Loop

If sTexto8 <> "" Then '*** PEAC 20110425
    sTexto8c = oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    sTexto8c = sTexto8c + "A S I E N T O   D E   C U E N T A   D E   O R D E N" & oImpresora.gPrnSaltoLinea
    sTexto8c = sTexto8c + String(pnColPage - 1, "=") & oImpresora.gPrnSaltoLinea
    sTexto8c = sTexto8c + CON + "Cuenta Contable                                           DEBE       HABER" & oImpresora.gPrnSaltoLinea
    sTexto8c = sTexto8c + String(pnColPage - 1, "-") & oImpresora.gPrnSaltoLinea
    sTexto8c = sTexto8c + sTexto8
    sTexto8c = sTexto8c + String(pnColPage - 1, "-") & oImpresora.gPrnSaltoLinea
    '''sTexto8c = sTexto8c + Space(75) & "TOTALES        " & gcMN & Right(Space(15) & Format(lnTotalDebe8, "###,###,##0.00"), 15) & "  " & gcMN & Right(Space(15) & Format(lnTotalHaber8, "###,###,##0.00"), 15) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    sTexto8c = sTexto8c + Space(75) & "TOTALES        " & gcPEN_SIMBOLO & Right(Space(15) & Format(lnTotalDebe8, "###,###,##0.00"), 15) & "  " & gcPEN_SIMBOLO & Right(Space(15) & Format(lnTotalHaber8, "###,###,##0.00"), 15) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    sTexto8c = sTexto8c + String(pnColPage - 1, "=") & oImpresora.gPrnSaltoLinea
End If

rs.Close
Set rs = Nothing
lsImp = sAsiento & oImpresora.gPrnSaltoLinea & ImpreAsiento(pnColPage, psSimbolo, sTexto, lnTotalDebe, lnTotalHaber, True, "ENCARGADO")

lsImp = lsImp + sTexto8c '*** PEAC 20110425

''Ahora Imprimimos Detalle de Documentos por Cuenta en el Debe
Set rs = oCh.GetDocumentosAsientoCH(psAreaCh, psAgeCh, pnProcCH, psCtaFondoFijo)
sCabe = CON + String(125, "=") & oImpresora.gPrnSaltoLinea _
      & "  D O C U M E N T O               C O N C E P T O                                                DEBE           HABER" & oImpresora.gPrnSaltoLinea _
      & "TIPO  NUMERO        FECHA " & oImpresora.gPrnSaltoLinea _
      & String(125, "-") & oImpresora.gPrnSaltoLinea

nLi = pnLinPage: P = 0
sTexto = PrnSet("MI", 4)
Do While Not rs.EOF
   If nLi > pnLinPage - 6 Then
      Linea sTexto, COFF + Cabecera("DETALLE DE GASTOS DE CAJA CHICA POR DOCUMENTO", P, psSimbolo, pnColPage, sCabe, , gsNomCmac)
      nLi = 9
   End If
   Linea sTexto, CON + BON + "CUENTA CONTABLE: " & ImpreFormat(rs!cCtaContCod, 18) & ImpreFormat(rs!cCtaContDesc, 80) + BOFF, 2
   nLi = nLi + 2
   sAsiento = rs!cCtaContCod
   lnTotalDebe = 0: lnTotalHaber = 0
   Do While sAsiento = rs!cCtaContCod
      lnTotalDebe = lnTotalDebe + rs!Debe
      lnTotalHaber = lnTotalHaber + rs!Haber
      Linea sTexto, ImpreFormat(rs!cDocAbrev, 3, 0) & ImpreFormat(rs!cDocNro, 16) & ImpreFormat(rs!dDocFecha, 10) & ImpreFormat(rs!cMovdesc, 50) & ImpreFormat(rs!Debe, 14, 2, True) & ImpreFormat(rs!Haber, 14, 2, True)
      nLi = nLi + 1
      If nLi > pnLinPage - 6 Then
         Linea sTexto, COFF + Cabecera("DETALLE DE GASTOS DE CAJA CHICA POR DOCUMENTO", P, psSimbolo, pnColPage, sCabe, , gsNomCmac)
         nLi = 9
      End If
      rs.MoveNext
      If rs.EOF Then Exit Do
   Loop
   Linea sTexto, String(125, "-")
   Linea sTexto, BON + ImpreFormat("", 40) & ImpreFormat("TOTAL", 40) & ImpreFormat(lnTotalDebe, 15, 2, True) & " " & ImpreFormat(lnTotalHaber, 13, 2, True) + BOFF, 2
   nLi = nLi + 3
Loop
lsImp = lsImp & oImpresora.gPrnSaltoPagina & CON & sTexto & COFF

ImprimeAsientoRendicionCH = lsImp


End Function
Public Function ImpreAsiento(ByVal pnColPage As Integer, ByVal psSimbolo As String, _
                            ByVal psAsiento As String, ByVal pnTotalDebe As Currency, _
                            ByVal pnTotalHaber As Currency, lPiePage As Boolean, Optional lsArea As String = "CAJA") As String
Dim sImpre As String, nLon  As String
sImpre = sImpre & oImpresora.gPrnSaltoLinea
' ASIENTO CONTABLE
'''sImpre = sImpre & BON & "ASIENTO CONTABLE EN " & IIf(psSimbolo = "S/.", "M.N. :", "M.E. :") & BOFF & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
sImpre = sImpre & BON & "ASIENTO CONTABLE EN " & IIf(psSimbolo = gcPEN_SIMBOLO, "M.N. :", "M.E. :") & BOFF & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
sImpre = sImpre & String(pnColPage - 1, "=") & oImpresora.gPrnSaltoLinea
sImpre = sImpre & "  " & BON & "Cuenta Contable" & Space(43) & "DEBE       HABER" & BOFF & oImpresora.gPrnSaltoLinea
sImpre = sImpre & String(pnColPage - 1, "-") & oImpresora.gPrnSaltoLinea
sImpre = sImpre & CON & psAsiento & COFF     'Detalle de Asiento
sImpre = sImpre & String(pnColPage - 1, "-") & oImpresora.gPrnSaltoLinea
sImpre = sImpre & CON
sImpre = sImpre & Space(75) & "TOTALES        " & psSimbolo & Right(Space(15) & Format(pnTotalDebe, "###,###,##0.00"), 15) & "   " & psSimbolo & Right(Space(15) & Format(pnTotalHaber, "###,###,##0.00"), 15) & oImpresora.gPrnSaltoLinea
sImpre = sImpre & COFF
sImpre = sImpre & String(pnColPage - 1, "=") & oImpresora.gPrnSaltoLinea

If lPiePage Then
   sImpre = sImpre & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
   sImpre = sImpre & BON
   sImpre = sImpre & Space(5) & "_____________________                           ______________________" & oImpresora.gPrnSaltoLinea
   lsArea = "Vo Bo " & lsArea
   nLon = Int((21 - Len(lsArea)) / 2)
   sImpre = sImpre & Space(5) & Space(nLon) & lsArea & Space(nLon) & "                              Vo Bo CONTABILIDAD" & BOFF & oImpresora.gPrnSaltoLinea
End If
ImpreAsiento = sImpre
End Function
Public Function ImprimeDocSalidaEfectivo(ByVal pnColPage As Integer, ByVal psFecha As String, _
                                        ByVal psOpecod As String, ByVal psMovNro As String, ByVal psNomCmact As String, Optional ByVal lnNumCop As Long = 1)
Dim lsImpre As String
Dim psTit1 As String
Dim rs As ADODB.Recordset
Dim oCaja As nCajaGeneral
Dim lsAreaOrigen As String
Dim lsAreaDestino As String
Dim lsTipoTransporte As String
Dim lsNombreTransp As String
Dim lnImporteTrans As Currency
Dim lsNroDocTransp As String
Dim lsNroDocSal As String
Dim lsBilletaje As String
Dim I As Integer
Set oCaja = New nCajaGeneral

Set rs = oCaja.GetDatosHabilitacion(psMovNro)
If Not rs.EOF And Not rs.BOF Then
    Do While Not rs.EOF
        If rs!Origen <> "" Then lsAreaOrigen = rs!Origen
        If rs!Destino <> "" Then lsAreaDestino = rs!Destino
        lsTipoTransporte = rs!cTipoTrans
        lsNombreTransp = Trim(rs!cNombreTransp)
        lnImporteTrans = rs!nMontoTrans
        lsNroDocTransp = rs!DocCompTransp
        lsNroDocSal = rs!DocSalEfecBov
        rs.MoveNext
    Loop
Else
    Exit Function
End If
rs.Close
Set rs = Nothing
lsImpre = ""
psTit1 = "COMPROBANTE DE SALIDA EN EFECTIVO  " & IIf(lsNroDocSal = "", "", " N° " & lsNroDocSal)
lsImpre = lsImpre & ImpreCabAsiento(pnColPage, CDate(psFecha), psNomCmact, psOpecod, psMovNro, psTit1, True) & oImpresora.gPrnSaltoLinea
Linea lsImpre, ImpreFormat("Area/Agencia Origen ", 25) & ":" & ImpreFormat(lsAreaOrigen, 40)
Linea lsImpre, ImpreFormat("Area/Agencia Destino", 25) & ":" & ImpreFormat(lsAreaDestino, 40), 2
Linea lsImpre, BON + ImpreFormat("TRANSPORTE", 25) + BOFF
If lsNombreTransp <> "" Then
    Linea lsImpre, ImpreFormat("TIPO ", 20, 5) & ":" & ImpreFormat(lsTipoTransporte, 40)
    Linea lsImpre, ImpreFormat("EMPRESA", 20, 5) & ":" & ImpreFormat(lsNombreTransp, 50)
    Linea lsImpre, ImpreFormat("DOCUMENTO", 20, 5) & ":" & ImpreFormat(IIf(lsNroDocTransp = "", "NO INDICADO!!!", lsNroDocTransp), 20)
    '''Linea lsImpre, BON + ImpreFormat("COSTO TRANSPORTE", 20, 5) & ":  S/." & ImpreFormat(lnImporteTrans, 10, 2, True) + BOFF, 3 'MARG ERS044-2016
    Linea lsImpre, BON + ImpreFormat("COSTO TRANSPORTE", 20, 5) & ":  " & gcPEN_SIMBOLO & ImpreFormat(lnImporteTrans, 10, 2, True) + BOFF, 3 'MARG ERS044-2016
Else
    Linea lsImpre, ImpreFormat("NO INDICADO", 20, 5), 3
End If
lsBilletaje = ImprimeBilletaje(Mid(psOpecod, 3, 1), psMovNro, pnColPage, False)
Linea lsImpre, lsBilletaje
'Linea lsImpre, ImprePiePag(pnColPage, "158", "Responsable Traslado")
Linea lsImpre, ImprePiePag(pnColPage, IIf(lsNroDocSal = "", "158", "1"), IIf(lsNroDocSal = "", "Responsable Traslado", ""))

Dim lsAux As String
For I = 1 To lnNumCop - 1
    lsAux = lsAux & oImpresora.gPrnSaltoPagina & lsImpre
Next
lsImpre = lsImpre & lsAux
ImprimeDocSalidaEfectivo = lsImpre

End Function
Public Function ImprimeDocConfHabEfectivo(ByVal pnColPage As Integer, ByVal psFecha As String, _
                                        ByVal psOpecod As String, ByVal psMovNro As String, ByVal psNomCmact As String)
Dim lsImpre As String
Dim psTit1 As String
Dim rs As ADODB.Recordset
Dim oCaja As nCajaGeneral
Dim lsAreaOrigen As String
Dim lsAreaDestino As String
Dim lsTipoTransporte As String
Dim lsNombreTransp As String
Dim lnImporteTrans As Currency
Dim lsNroDocTransp As String
Dim lsNroDocSal As String
Dim lsBilletaje As String

Set oCaja = New nCajaGeneral

Set rs = oCaja.GetDatosConfHabAgencias(psMovNro)
If Not rs.EOF And Not rs.BOF Then
    Do While Not rs.EOF
        If rs!Origen <> "" Then lsAreaOrigen = rs!Origen
        If rs!Destino <> "" Then lsAreaDestino = rs!Destino
        lsTipoTransporte = rs!cTipoTrans
        lsNombreTransp = Trim(rs!cNombreTransp)
        lnImporteTrans = rs!nMontoTrans
        lsNroDocTransp = rs!DocCompTransp
        lsNroDocSal = rs!DocSalEfecBov
        rs.MoveNext
    Loop
Else
    Exit Function
End If
rs.Close
Set rs = Nothing
lsImpre = ""
psTit1 = "CONFIRMACION DE HABILITACION EN EFECTIVO"
lsImpre = lsImpre & ImpreCabAsiento(pnColPage, CDate(psFecha), psNomCmact, psOpecod, psMovNro, psTit1, True) & oImpresora.gPrnSaltoLinea
Linea lsImpre, ImpreFormat("Area/Agencia Origen ", 25) & ":" & ImpreFormat(lsAreaOrigen, 40)
Linea lsImpre, ImpreFormat("Area/Agencia Destino", 25) & ":" & ImpreFormat(lsAreaDestino, 40), 3

'Linea lsImpre, BON + ImpreFormat("TRANSPORTE", 25) + BOFF
'If lsNombreTransp <> "" Then
'    Linea lsImpre, ImpreFormat("TIPO ", 20, 5) & ":" & ImpreFormat(lsTipoTransporte, 40)
'    Linea lsImpre, ImpreFormat("EMPRESA", 20, 5) & ":" & ImpreFormat(lsNombreTransp, 50)
'    Linea lsImpre, ImpreFormat("DOCUMENTO", 20, 5) & ":" & ImpreFormat(IIf(lsNroDocTransp = "", "NO INDICADO!!!", lsNroDocTransp), 20)
'    Linea lsImpre, BON + ImpreFormat("COSTO TRANSPORTE", 20, 5) & ":  S/." & ImpreFormat(lnImporteTrans, 10, 2, True) + BOFF, 3
'Else
'    Linea lsImpre, ImpreFormat("NO INDICADO", 20, 5), 3
'End If
lsBilletaje = ImprimeBilletaje(Mid(psOpecod, 3, 1), psMovNro, pnColPage, False)
Linea lsImpre, lsBilletaje
'Linea lsImpre, ImprePiePag(pnColPage, "158", "Responsable Traslado")
Linea lsImpre, ImprePiePag(pnColPage, "5")
ImprimeDocConfHabEfectivo = lsImpre
End Function
' Procedimeinto que imprime la boleta de Operación
' luego de una operacion de cajero por impresora
Public Sub ImprimeBoletahabilitacion(ByVal psTitCab As String, ByVal psTitDet As String, _
                                    ByVal psUserOrig As String, ByVal psNomUserOrigen As String, _
                                    ByVal psUserDest As String, ByVal psNomUserDest As String, _
                                    ByVal pnMoneda As Moneda, ByVal psCodOpe As String, ByVal pnImporte As String, _
                                    ByVal psNomAge As String, _
                                    ByVal psMovNro As String, ByVal psLpt As String)

Dim nFicSal As Integer
Dim lsFecha As String
Dim lsHora As String
Dim lsUserOrig As String
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lsMensaje As String * 39
Dim I As Integer
Dim lnNumLineas As Integer
Dim lnLineasBol As Integer
Dim lnEspaceIni As Integer
Dim lnAnchoBol As Long

ETIQ:
On Error GoTo ERROR

'lnTope = 0 '6 'Tope de lineas en Boleta

lsNegritaOn = oImpresora.gPrnBoldON
lsNegritaOff = oImpresora.gPrnBoldOFF

nFicSal = FreeFile
Open psLpt For Output As nFicSal

Print #nFicSal, oImpresora.gPrnInicializa;

Print #nFicSal, oImpresora.gPrnEspaLineaN;   'espaciamiento lineas 1/6 pulg.
Print #nFicSal, oImpresora.gPrnEspaLineaValor & Chr$(lnLineasBol);   'Longitud de página a 22 líneas'
Print #nFicSal, oImpresora.gPrnTamLetra10CPI;    'Tamaño 10 cpi
Print #nFicSal, oImpresora.gPrnTpoLetraRoman;      'Tipo de Letra Sans Serif
Print #nFicSal, oImpresora.gPrnCondensadaOFF  ' cancela condensada
Print #nFicSal, oImpresora.gPrnBoldOFF  ' desactiva negrita

lsFecha = Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4)
'20010703 1347191120100EJRS
lsHora = Format$(Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2), "hh:mm:ss")
lsUserOrig = Right(psMovNro, 4)
lnEspaceIni = 2
lnLineasBol = 22
lnAnchoBol = 50

Print #nFicSal, ImpreCarEsp(PrnSet("B+") + ImpreFormat("CMACT - OPERACIONES", lnAnchoBol, lnEspaceIni) & ImpreFormat("CMACT - OPERACIONES", lnAnchoBol, lnEspaceIni))
'''Print #nFicSal, ImpreCarEsp(ImpreFormat(Trim(psNomAge) + IIf(pnMoneda = gMonedaNacional, "-SOLES", "-DOLARES"), lnAnchoBol, lnEspaceIni) & ImpreFormat(Trim(psNomAge) + IIf(pnMoneda = gMonedaNacional, "-SOLES", "-DOLARES"), lnAnchoBol, lnEspaceIni) + PrnSet("B-")) 'MARG ERS044-2016
Print #nFicSal, ImpreCarEsp(ImpreFormat(Trim(psNomAge) + IIf(pnMoneda = gMonedaNacional, "-" & StrConv(gcPEN_PLURAL, vbUpperCase), "-DOLARES"), lnAnchoBol, lnEspaceIni) & ImpreFormat(Trim(psNomAge) + IIf(pnMoneda = gMonedaNacional, "-" & StrConv(gcPEN_PLURAL, vbUpperCase), "-DOLARES"), lnAnchoBol, lnEspaceIni) + PrnSet("B-")) 'MARG ERS044-2016
Print #nFicSal, ImpreCarEsp(ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni))
Print #nFicSal, ""
Print #nFicSal, ImpreCarEsp(ImpreFormat("PARA :[" & psUserDest & "] " & psNomUserDest, lnAnchoBol, lnEspaceIni) & ImpreFormat("PARA :[" & psUserDest & "] " & psNomUserDest, lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat("DE   :[" & psUserOrig & "] " & psNomUserOrigen, lnAnchoBol, lnEspaceIni) & ImpreFormat("DE   :[" & psUserOrig & "] " & psNomUserOrigen, lnAnchoBol, lnEspaceIni))
Print #nFicSal, ""
'Print #nFicSal, ImpreCarEsp(PrnSet("B+") + CentrarCadena(psTitCab, lnAnchoBol - 10) & CentrarCadena(psTitCab, lnAnchoBol))
Print #nFicSal, ImpreCarEsp(PrnSet("B+") + Space(lnEspaceIni) + CentrarCadena(psTitCab, lnAnchoBol - 3) & CentrarCadena(psTitCab, lnAnchoBol - 3, lnEspaceIni + 4))
Print #nFicSal, ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni) & ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni) + PrnSet("B-"))
Print #nFicSal, ImpreCarEsp(ImpreFormat(ImpreFormat(psTitDet, lnAnchoBol / 2) & ImpreFormat(CCur(pnImporte), 15, 2, True), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat(psTitDet, lnAnchoBol / 2) & ImpreFormat(CCur(pnImporte), 15, 2, True), lnAnchoBol, lnEspaceIni))
For I = 1 To 3
    Print #nFicSal, ""
Next
Print #nFicSal, ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni) & ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat(lsUserOrig, lnAnchoBol, lnEspaceIni) & ImpreFormat(lsUserOrig, lnAnchoBol, lnEspaceIni))

Dim oGen As DGeneral
Set oGen = New DGeneral
lsMensaje = oGen.GetMensajeBoletas("")
Set oGen = Nothing
Print #nFicSal, PrnSet("B+") + ImpreCarEsp(ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) & ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) + PrnSet("B-"))
Print #nFicSal, ""
Print #nFicSal, ""

Close nFicSal
Exit Sub
ERROR:
    Close nFicSal
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ImprimeOperaciones() As String
Dim Rope As New ADODB.Recordset
Dim Robj As New ADODB.Recordset
Dim Rcta As New ADODB.Recordset
Dim Rdoc As New ADODB.Recordset
Dim sOpe As String, nOPE As Long
Dim sCta As String, nCta As Integer
Dim sObj As String, nObj As Integer
Dim sDoc As String, nDoc As Integer
Dim sSQL As String, nLin As Long, nPag As Integer
Dim I As Integer
Dim clsOpe As New DOperacion
Dim clsCta As New DCtaCont
Dim sTexto As String
Dim mMat(5000, 12) As String

On Error GoTo ImprimeOperacionesErr


Linea sTexto, ""
'Set Rope = clsOpe.CargaOpeTpo("")
Set Rope = clsOpe.CargaOpeTpo1("")
If Not Rope.EOF Then
   nOPE = 0
   RaiseEvent BarraShow(Rope.RecordCount)
   Do While Not Rope.EOF
      DoEvents
      RaiseEvent BarraProgress(Rope.Bookmark, "OPERACIONES", "Generando Reporte", "", vbBlue)
      nOPE = nOPE + 1
      mMat(nOPE, 0) = Format(Rope!nOpeNiv, "000")
      mMat(nOPE, 1) = IIf(Rope!cOpeCod = "", "", Rope!cOpeCod)
      mMat(nOPE, 2) = IIf(IsNull(Rope!cOpeDesc), "", Rope!cOpeDesc)
      sOpe = Rope!cOpeCod
      
      Set Rdoc = clsOpe.CargaOpeDoc(sOpe, 0)
      
      nDoc = nOPE
      If Not Rdoc.EOF Then
         Do While Not Rdoc.EOF
            mMat(nDoc, 10) = Rdoc!nDocTpo
            mMat(nDoc, 11) = Rdoc!cDocDesc
            nDoc = nDoc + 1
            Rdoc.MoveNext
         Loop
      End If
             
     Set Rcta = clsOpe.CargaOpeCta(sOpe)
      
      nCta = nOPE
      If Not Rcta.EOF Then
         Do While Not Rcta.EOF
            sCta = Rcta!cCtaContCod
            mMat(nCta, 3) = Rcta!cCtaContCod
            mMat(nCta, 4) = Rcta!cCtaContDesc
            mMat(nCta, 5) = IIf(Rcta!cOpeCtaDH = "D", "DEBE ", "HABER")
            mMat(nCta, 6) = IIf(Rcta!cOpeCtaOpc = "1", "Obligatoria", "Opcional")
            
            Set Robj = clsCta.CargaCtaObj(sCta, , True)
            If Not Robj.EOF Then
               nObj = nCta
               Do While Not Robj.EOF
                  mMat(nObj, 7) = IIf(Robj!cObjetoCod = "", "", Robj!cObjetoCod)
                  mMat(nObj, 8) = IIf(IsNull(Robj!cObjetoDesc), "", Robj!cObjetoDesc)
                  nObj = nObj + 1
                  Robj.MoveNext
               Loop
            End If
            
            nCta = nObj
            Rcta.MoveNext
         Loop
      End If
      If nDoc > nCta Then
         nOPE = nDoc
      Else
         nOPE = nCta
      End If
      If nOPE > 4990 Then Exit Do
      Rope.MoveNext
   Loop
End If

nLin = 1
nPag = 1
sTexto = ""
For I = 1 To nOPE
    DoEvents
    If nLin = 1 Then
       Linea sTexto, ImprimeOperacionesCab(nPag), 0
       nLin = nLin + 8
    End If
    If Val(mMat(I, 0)) = 1 Then
       Linea sTexto, Justifica(mMat(I, 1), 6) & " " & mMat(I, 2)
       Linea sTexto, String(Len(Justifica(mMat(I, 1), 6) & " " & mMat(I, 2)), "-")
       nLin = nLin + 1
    Else
       Linea sTexto, Justifica(mMat(I, 1), 6) & " " & Justifica(mMat(I, 2), 40) & " " & Justifica(mMat(I, 3), 12) & " " & Justifica(mMat(I, 4), 40) & " " & Justifica(mMat(I, 5), 5) & " " & Justifica(mMat(I, 6), 12) & " " & Justifica(mMat(I, 7), 20) & " " & Justifica(mMat(I, 8), 40) & "     " & " " & Justifica(mMat(I, 10), 4) & "  " & " " & Justifica(mMat(I, 11), 40)
    End If
    nLin = nLin + 1
    If nLin >= 60 Then
       nPag = nPag + 1
       Linea sTexto, String(215, "-") + oImpresora.gPrnSaltoPagina
       nLin = 1
    End If
Next
If Not sTexto = "" Then
   RaiseEvent BarraClose
End If
Set clsCta = Nothing
Set clsOpe = Nothing
RSClose Robj
RSClose Rope
RSClose Rcta
ImprimeOperaciones = sTexto
Exit Function
ImprimeOperacionesErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeOperaciones Method")
End Function

Private Function ImprimeOperacionesCab(nPag As Integer) As String
Dim lsTexto As String
Linea lsTexto, Justifica("CMACT", 15) + CentrarCadena("Pagina Nro " + Format(nPag, "00"), 185) + CStr(Date)
Linea lsTexto, "Sede Principal " + Space(185) + CStr(Time)
Linea lsTexto, CentrarCadena("Listado de Operaciones", 215)
Linea lsTexto, String(215, "=")
Linea lsTexto, "            O P E R A C I O N                                     C U E N T A   C O N T A B L E                                       O B J E T O S                                             D O C U M E N T O S"
Linea lsTexto, "---------------------------------------------   ---------------------------------------------------------------------    ---------------------------------------------------------------   ----------------------------"
Linea lsTexto, "CODIGO DESCRIPCION                              CUENTA      DESCRIPCION                            CLASE  TIPO           CODIGO               DESCRIPCION                                  TIPO    DENOMINACION  "
Linea lsTexto, String(215, "-")
ImprimeOperacionesCab = lsTexto
End Function

Public Function ImprimePlanContable(prs As Recordset, psNomTabla As String, pnLinPage As Integer) As String
' Variables de Sangrado de Cta
Dim nLongC, nLongT As Integer, nLongD As Integer, nNivel As Integer
Dim aSpace(1 To 20, 1 To 2) As Integer, nPos As Integer, Espacios As Integer
' Variables de Impresión
Dim lnCarLin, lnPagNro As Integer
Dim lsTitRep As String, lsCabe01 As String, lsCabe02 As String
Dim lnPiePag, lnLinea As Integer
Dim sTexto As String

For nPos = 1 To 20
    aSpace(nPos, 1) = 0
Next
nNivel = 0
nLongC = 0

lnCarLin = 90: lnLinea = 70
lnPiePag = 5: lnPagNro = 0

   If Trim(psNomTabla) = "CtaCont" Then
      lsTitRep = "P L A N   C O N T A B L E   G E N E R A L"
   Else
      lsTitRep = "P L A N   C O N T A B L E   B A S E   S B S"
   End If
  lsTitRep = String(Int((lnCarLin - Len(lsTitRep)) / 2), " ") & lsTitRep
sTexto = ""
Do While Not prs.EOF
   DoEvents
   ' Sangra Código
   nLongT = Len(Trim(prs!cCtaContCod))
'   If nLongC < nLongT Then
'      Espacios = Espacios + nLongT - 1
'      aSpace(nLongT, 1) = nLongT
'      aSpace(nLongT, 2) = Espacios
'      If nLongT <= 6 Then
'         If lnPagNro > 0 Then
'            Linea sTexto, ""
'            lnLinea = lnLinea + 1
'         End If
'      End If
'   End If
'   If nLongC > nLongT Then
'      Espacios = aSpace(nLongT, 2)
'      nNivel = nLongT
'      If nNivel <= 6 And lnPagNro > 0 Then
'         Linea sTexto, ""
'         lnLinea = lnLinea + 1
'      End If
'   End If
   Espacios = nLongT
   nLongC = Len(Trim(prs!cCtaContCod))
   nLongD = Len(Trim(prs!cCtaContDesc))
   
   'Salto de Pagina si cumple condicion
   If lnLinea > pnLinPage - lnPiePag - 1 Then
        If lnPagNro > 0 Then
           Linea sTexto, String(lnCarLin, "-") & oImpresora.gPrnSaltoPagina
        End If
        lnPagNro = lnPagNro + 1
        ' Definición de Cabecera 1
        lsCabe01 = vsEmpNom
        lsCabe01 = lsCabe01 & Space((lnCarLin - 34) - Len(lsCabe01))
        lsCabe01 = lsCabe01 & "PAGINA: " & FillNum(Trim(Str(lnPagNro)), 4, " ")
        lsCabe01 = lsCabe01 & Space(5) & "FECHA: " & Format(vsFecSis, gsFormatoFechaView)
        ' Definición de Cabecera 2
        lsCabe02 = "CONTABILIDAD" & "-" & vsAgeNom
        lsCabe02 = lsCabe02 & Space((lnCarLin - 17) - Len(lsCabe02))
        lsCabe02 = lsCabe02 & "HORA :   " & Format(Now(), "hh:mm:ss")
        
   '    Cabecera de Impresión
        Linea sTexto, lsCabe01
        Linea sTexto, lsCabe02

        Linea sTexto, UCase(lsTitRep), 2
        Linea sTexto, String(lnCarLin, "-")
        Linea sTexto, "   CODIGO   " & String(15, " ") & "DESCRIPCION" & String(12, " ") & "   "
        Linea sTexto, String(lnCarLin, "-"), 2
        lnLinea = 9
    End If
    Linea sTexto, Space(Espacios) & Trim(prs!cCtaContCod) _
        & " " & Left(prs!cCtaContDesc, nLongD + 1)
    lnLinea = lnLinea + 1
    prs.MoveNext
Loop
Linea sTexto, String(lnCarLin, "-")             ' hasta 160 caracteres x línea
ImprimePlanContable = sTexto
End Function

Public Function ImprimeObjetos(pnLinPage As Integer) As String
Dim oObj As DObjeto
Dim prs  As ADODB.Recordset

Set oObj = New DObjeto
   Set prs = oObj.CargaObjeto(, , adLockReadOnly)
Set oObj = Nothing

' Variables de Impresión
Dim lnCarLin As Integer, lnPageNro As Integer
Dim lnLinea As Integer
Dim lsImpre As String

lnCarLin = pnLinPage: lnLinea = pnLinPage
lnPageNro = 0

Do While Not prs.EOF
   If lnLinea > lnCarLin - 2 Then
      Linea lsImpre, Cabecera("REPORTE DE OBJETOS", lnPageNro, "", pnLinPage, , , gsNomCmac), 0
      Linea lsImpre, " CODIGO        DESCRIPCION"
      Linea lsImpre, String(79, "-")
      lnLinea = 6
   End If
   Linea lsImpre, Space(2) & Trim(prs!cObjetoCod) _
               & Space(5) & Trim(prs!cObjetoDesc)
   lnLinea = lnLinea + 1
   prs.MoveNext
Loop
prs.Close: Set prs = Nothing
Linea lsImpre, String(lnCarLin, "-")
ImprimeObjetos = lsImpre
End Function
Public Sub ImprimeBoletaCompraVenta(ByVal psTitCab As String, ByVal psTitDet As String, _
                                    ByVal psNomPers As String, ByVal psDirPers As String, _
                                    ByVal psDocumento As String, ByVal pnTipoCambio As Currency, _
                                    ByVal psCodOpe As String, ByVal pnMontoDol As Currency, ByVal pnMontoSoles As Currency, _
                                    ByVal psNomAge As String, ByVal psMovNro As String, ByVal psLpt As String)

Dim nFicSal As Integer
Dim lsFecha As String
Dim lsHora As String
Dim lsUserOrig As String
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lsMensaje As String * 39
Dim I As Integer
Dim lnNumLineas As Integer
Dim lnLineasBol As Integer
Dim lnEspaceIni As Integer
Dim lnAnchoBol As Long
Dim psSubTit As String
Dim psSubTit2 As String
Dim lnMovNro As Long
Dim oMov As DMov
Set oMov = New DMov
lnMovNro = oMov.GetnMovNro(psMovNro)
Set oMov = Nothing

ETIQ:
On Error GoTo ERROR

lsNegritaOn = oImpresora.gPrnBoldON
lsNegritaOff = oImpresora.gPrnBoldOFF

nFicSal = FreeFile
Open psLpt For Output As nFicSal

Print #nFicSal, oImpresora.gPrnInicializa;

Print #nFicSal, oImpresora.gPrnEspaLineaN;   'espaciamiento lineas 1/6 pulg.
Print #nFicSal, oImpresora.gPrnEspaLineaValor + Chr$(lnLineasBol);   'Longitud de página a 22 líneas'
Print #nFicSal, oImpresora.gPrnTamLetra10CPI;    'Tamaño 10 cpi
Print #nFicSal, oImpresora.gPrnTpoLetraRoman;      'Tipo de Letra Sans Serif
Print #nFicSal, oImpresora.gPrnCondensadaOFF  ' cancela condensada
Print #nFicSal, oImpresora.gPrnBoldOFF  ' desactiva negrita

Select Case psCodOpe
    Case gOpeCajeroMECompra
        psTitDet = "Monto Comprado $"
        '''psSubTit = "Neto a Pagar S/." 'MARG ERS044-2016
        psSubTit = "Neto a Pagar " & gcPEN_SIMBOLO 'MARG ERS044-2016
        psSubTit2 = "Tipo de Cambio Compra"
    Case gOpeCajeroMEVenta
        psTitDet = "Monto Vendido $"
        '''psSubTit = "Neto a Recibir S/." 'MARG ERS044-2016
        psSubTit = "Neto a Recibir " & gcPEN_SIMBOLO 'MARG ERS044-2016
        psSubTit2 = "Tipo de Cambio Venta"
End Select

lsFecha = Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4)
'20010703 1347191120100EJRS
lsHora = Format$(Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2), "hh:mm:ss")
lsUserOrig = Right(psMovNro, 4)
lnEspaceIni = 2
lnLineasBol = 22
lnAnchoBol = 50

Print #nFicSal, ImpreCarEsp(PrnSet("B+") + ImpreFormat("CMACT - OPERACIONES", lnAnchoBol, lnEspaceIni) & ImpreFormat("CMACT - OPERACIONES", lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat(ImpreFormat(psNomAge, 30) & ImpreFormat(lnMovNro, 10, 0, False), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat(psNomAge, 30) & ImpreFormat(lnMovNro, 10, 0, False), lnAnchoBol, lnEspaceIni) + PrnSet("B-"))
Print #nFicSal, ImpreCarEsp(ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat(psNomPers, lnAnchoBol, lnEspaceIni) & ImpreFormat(psNomPers, lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat("Dir:" & psDirPers, lnAnchoBol, lnEspaceIni) & ImpreFormat("Dir:" & psDirPers, lnAnchoBol, lnEspaceIni))
If psDocumento <> "" Then
    Print #nFicSal, ImpreCarEsp(ImpreFormat("Nro Doc:" & psDocumento, lnAnchoBol, lnEspaceIni) & ImpreFormat("Nro Doc:" + psDocumento, lnAnchoBol, lnEspaceIni))
Else
    Print #nFicSal, ""
End If
Print #nFicSal, ImpreCarEsp(PrnSet("B+") + Space(lnEspaceIni) + CentrarCadena(psTitCab, lnAnchoBol - 3) & CentrarCadena(psTitCab, lnAnchoBol - 3, lnEspaceIni + 4))
Print #nFicSal, ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni) & ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat(ImpreFormat(psTitDet, lnAnchoBol / 2) & ImpreFormat(CCur(pnMontoDol), 15, 2, True), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat(psTitDet, lnAnchoBol / 2) & ImpreFormat(CCur(pnMontoDol), 15, 2, True), lnAnchoBol, lnEspaceIni) + PrnSet("B-"))
Print #nFicSal, ""
Print #nFicSal, ""
Print #nFicSal, ImpreCarEsp(ImpreFormat(ImpreFormat(psSubTit, lnAnchoBol / 2) & ImpreFormat(CCur(pnMontoSoles), 15, 2, True), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat(psSubTit, lnAnchoBol / 2) & ImpreFormat(CCur(pnMontoSoles), 15, 2, True), lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat(ImpreFormat(psSubTit2, lnAnchoBol / 2) & ImpreFormat(CCur(pnTipoCambio), 15, 2, True), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat(psSubTit2, lnAnchoBol / 2) & ImpreFormat(CCur(pnTipoCambio), 15, 2, True), lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni) & ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat(Right(psMovNro, 4), lnAnchoBol, lnEspaceIni) & ImpreFormat(Right(psMovNro, 4), lnAnchoBol, lnEspaceIni))

Dim oGen As DGeneral
Set oGen = New DGeneral
lsMensaje = oGen.GetMensajeBoletas("")
Set oGen = Nothing
Print #nFicSal, PrnSet("B+") + ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) & ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) + PrnSet("B-")
Print #nFicSal, ""
Print #nFicSal, ""
Print #nFicSal, ""

Close nFicSal
Exit Sub
ERROR:
    Close nFicSal
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub ImprimeBoletaExtornos(ByVal psTitCab As String, ByVal psTitDet As String, _
                                 ByVal psCodOpe As String, ByVal pnImporte As Currency, _
                                 ByVal psUserOpe As String, ByVal psNomPers As String, _
                                 ByVal psNomAge As String, ByVal psMovNro As String, ByVal psLpt As String)

Dim nFicSal As Integer
Dim lsFecha As String
Dim lsHora As String
Dim lsUserOrig As String
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lsMensaje As String * 39
Dim I As Integer
Dim lnNumLineas As Integer
Dim lnLineasBol As Integer
Dim lnEspaceIni As Integer
Dim lnAnchoBol As Long
Dim psSubTit As String
Dim psSubTit2 As String
Dim lnMovNro As Long
Dim oMov As DMov
Dim lsMoneda As String
Set oMov = New DMov
lnMovNro = oMov.GetnMovNro(psMovNro)
Set oMov = Nothing

ETIQ:
On Error GoTo ERROR

lsNegritaOn = oImpresora.gPrnBoldON
lsNegritaOff = oImpresora.gPrnBoldOFF

nFicSal = FreeFile
Open psLpt For Output As nFicSal

Print #nFicSal, oImpresora.gPrnInicializa;

Print #nFicSal, oImpresora.gPrnEspaLineaN;   'espaciamiento lineas 1/6 pulg.
Print #nFicSal, oImpresora.gPrnEspaLineaValor + Chr$(lnLineasBol);   'Longitud de página a 22 líneas'
Print #nFicSal, oImpresora.gPrnTamLetra10CPI;    'Tamaño 10 cpi
Print #nFicSal, oImpresora.gPrnTpoLetraRoman;      'Tipo de Letra Sans Serif
Print #nFicSal, oImpresora.gPrnCondensadaOFF  ' cancela condensada
Print #nFicSal, oImpresora.gPrnBoldOFF  ' desactiva negrita

lsFecha = Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4)
'20010703 1347191120100EJRS
lsHora = Format$(Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2), "hh:mm:ss")
lsUserOrig = Right(psMovNro, 4)
lnEspaceIni = 2
lnLineasBol = 22
lnAnchoBol = 50

If Mid(psCodOpe, 3, 1) <> gMonedaNacional And Mid(psCodOpe, 3, 1) <> gMonedaExtranjera Then
    lsMoneda = ""
Else
    '''lsMoneda = IIf(Mid(psCodOpe, 3, 1) = gMonedaNacional, "-SOLES", "-DOLARES") 'MARG ERS044-2016
    lsMoneda = IIf(Mid(psCodOpe, 3, 1) = gMonedaNacional, "-" & StrConv(gcPEN_PLURAL, vbUpperCase), "-DOLARES") 'MARG ERS044-2016
End If

Print #nFicSal, ImpreCarEsp(PrnSet("B+") + ImpreFormat("CMACT - OPERACIONES", lnAnchoBol, lnEspaceIni) & ImpreFormat("CMACT - OPERACIONES", lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat(ImpreFormat(psNomAge, 20) & ImpreFormat(lsMoneda, 10) & ImpreFormat(lnMovNro, 10, 0, False), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat(psNomAge, 20) & ImpreFormat(lsMoneda, 10) & ImpreFormat(lnMovNro, 10, 0, False), lnAnchoBol, lnEspaceIni) + PrnSet("B-"))
Print #nFicSal, ImpreCarEsp(ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat(psNomPers, lnAnchoBol, lnEspaceIni) & ImpreFormat(psNomPers, lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat("Operación de :  " & psUserOpe, lnAnchoBol, lnEspaceIni) & ImpreFormat("Operación de :  " & psUserOpe, lnAnchoBol, lnEspaceIni))
Print #nFicSal, ""
'Print #nFicSal, ImpreCarEsp(PrnSet("B+") + CentrarCadena(psTitCab, lnAnchoBol - 3) + CentrarCadena(psTitCab, lnAnchoBol - 3, lnEspaceIni + 3))
Print #nFicSal, ImpreCarEsp(PrnSet("B+") + Space(lnEspaceIni) + CentrarCadena(psTitCab, lnAnchoBol - 3) & CentrarCadena(psTitCab, lnAnchoBol - 3, lnEspaceIni + 4))
Print #nFicSal, ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni) & ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat(ImpreFormat("MONTO EXTORNADO ", lnAnchoBol / 2) & ImpreFormat(CCur(pnImporte), 15, 2, True), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat("MONTO EXTORNADO", lnAnchoBol / 2) & ImpreFormat(CCur(pnImporte), 15, 2, True), lnAnchoBol, lnEspaceIni) + PrnSet("B-"))
Print #nFicSal, ""
Print #nFicSal, ""
Print #nFicSal, ImpreCarEsp(PrnSet("C+") + ImpreFormat(psTitDet, lnAnchoBol + 25, lnEspaceIni + 2) & ImpreFormat(psTitDet, lnAnchoBol + 25, lnEspaceIni + 10) + PrnSet("C-"))
Print #nFicSal, ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni) & ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat(Right(psMovNro, 4), lnAnchoBol, lnEspaceIni) & ImpreFormat(Right(psMovNro, 4), lnAnchoBol, lnEspaceIni))

Dim oGen As DGeneral
Set oGen = New DGeneral
lsMensaje = oGen.GetMensajeBoletas("")
Set oGen = Nothing
Print #nFicSal, PrnSet("B+") + ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) & ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) + PrnSet("B-")
Print #nFicSal, ""
Print #nFicSal, ""
Print #nFicSal, ""

Close nFicSal
Exit Sub
ERROR:
    Close nFicSal
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub ImprimeBoletaGeneral(ByVal psTitCab As String, ByVal psTitDet As String, _
                                 ByVal psCodOpe As String, ByVal pnImporte As Currency, _
                                 ByVal psNomAge As String, ByVal psMovNro As String, _
                                 ByVal psLpt As String, _
                                 Optional ByVal psNomPers As String = "", _
                                 Optional ByVal psNroDoc As String = "", _
                                 Optional ByVal psReferencia As String = "", _
                                 Optional ByVal psGlosa As String = "")

Dim nFicSal As Integer
Dim lsFecha As String
Dim lsHora As String
Dim lsUserOrig As String
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lsMensaje As String * 39
Dim I As Integer
Dim lnNumLineas As Integer
Dim lnLineasBol As Integer
Dim lnEspaceIni As Integer
Dim lnAnchoBol As Long
Dim psSubTit As String
Dim psSubTit2 As String
Dim lnMovNro As Long
Dim oMov As DMov
Dim lsMoneda As String
Set oMov = New DMov
lnMovNro = oMov.GetnMovNro(psMovNro)
Set oMov = Nothing

ETIQ:
On Error GoTo ERROR

lsNegritaOn = oImpresora.gPrnBoldON
lsNegritaOff = oImpresora.gPrnBoldOFF

nFicSal = FreeFile
Open psLpt For Output As nFicSal

Print #nFicSal, oImpresora.gPrnInicializa;

Print #nFicSal, oImpresora.gPrnEspaLineaN;   'espaciamiento lineas 1/6 pulg.
Print #nFicSal, oImpresora.gPrnEspaLineaValor + Chr$(lnLineasBol);   'Longitud de página a 22 líneas'
Print #nFicSal, oImpresora.gPrnTamLetra10CPI;    'Tamaño 10 cpi
Print #nFicSal, oImpresora.gPrnTpoLetraRoman;      'Tipo de Letra Sans Serif
Print #nFicSal, oImpresora.gPrnCondensadaOFF  ' cancela condensada
Print #nFicSal, oImpresora.gPrnBoldOFF  ' desactiva negrita

lsFecha = Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4)
'20010703 1347191120100EJRS
lsHora = Format$(Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2), "hh:mm:ss")
lsUserOrig = Right(psMovNro, 4)
lnEspaceIni = 2
lnLineasBol = 22
lnAnchoBol = 50

If Mid(psCodOpe, 3, 1) <> gMonedaNacional And Mid(psCodOpe, 3, 1) <> gMonedaExtranjera Then
    '''lsMoneda = "-SOLES" 'MARG ERS044-2016
    lsMoneda = "-" & StrConv(gcPEN_PLURAL, vbUpperCase)
Else
    '''lsMoneda = IIf(Mid(psCodOpe, 3, 1) = gMonedaNacional, "-SOLES", "-DOLARES") 'MARG ERS044-2016
    lsMoneda = IIf(Mid(psCodOpe, 3, 1) = gMonedaNacional, "-" & StrConv(gcPEN_PLURAL, vbUpperCase), "-DOLARES") 'MARG ERS044-2016
End If

Print #nFicSal, ImpreCarEsp(PrnSet("B+") + ImpreFormat("CMACT - OPERACIONES", lnAnchoBol, lnEspaceIni) & ImpreFormat("CMACT - OPERACIONES", lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat(ImpreFormat(psNomAge, 20) & ImpreFormat(lsMoneda, 10) & ImpreFormat(lnMovNro, 10, 0, False), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat(psNomAge, 20) & ImpreFormat(lsMoneda, 10) & ImpreFormat(lnMovNro, 10, 0, False), lnAnchoBol, lnEspaceIni) + PrnSet("B-"))
Print #nFicSal, ImpreCarEsp(ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat(psNomPers, lnAnchoBol, lnEspaceIni) & ImpreFormat(psNomPers, lnAnchoBol, lnEspaceIni))
Print #nFicSal, ""
Print #nFicSal, ""
Print #nFicSal, ImpreCarEsp(PrnSet("B+") + Space(lnEspaceIni) + CentrarCadena(psTitCab, lnAnchoBol - 3) & CentrarCadena(psTitCab, lnAnchoBol - 3, lnEspaceIni + 4))
Print #nFicSal, ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni) & ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat(ImpreFormat(psTitDet, lnAnchoBol / 2) & ImpreFormat(CCur(pnImporte), 15, 2, True), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat(psTitDet, lnAnchoBol / 2) & ImpreFormat(CCur(pnImporte), 15, 2, True), lnAnchoBol, lnEspaceIni) + PrnSet("B-"))
Print #nFicSal, ""
Print #nFicSal, ""
If psNroDoc <> "" Then
    Print #nFicSal, ImpreCarEsp(ImpreFormat(ImpreFormat("N° Documento :", lnAnchoBol / 2) & ImpreFormat(psNroDoc, 15), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat("N° Documento :", lnAnchoBol / 2) & ImpreFormat(psNroDoc, 15), lnAnchoBol, lnEspaceIni) + PrnSet("B-"))
End If
If psReferencia <> "" Then
    Print #nFicSal, ImpreCarEsp(ImpreFormat(ImpreFormat("Referencia :", lnAnchoBol / 2) & ImpreFormat(psReferencia, 15), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat("Referencia", lnAnchoBol / 2) & ImpreFormat(psReferencia, 15), lnAnchoBol, lnEspaceIni) + PrnSet("B-"))
End If
If psGlosa <> "" Then
    Print #nFicSal, ImpreCarEsp(PrnSet("C+") + ImpreFormat(psGlosa, lnAnchoBol + 25, lnEspaceIni + 2) & ImpreFormat(psGlosa, lnAnchoBol + 25, lnEspaceIni + 10) + PrnSet("C-"))
End If
Print #nFicSal, ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni) & ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni))
Print #nFicSal, ImpreCarEsp(ImpreFormat(Right(psMovNro, 4), lnAnchoBol, lnEspaceIni) & ImpreFormat(Right(psMovNro, 4), lnAnchoBol, lnEspaceIni))

Dim oGen As DGeneral
Set oGen = New DGeneral
lsMensaje = oGen.GetMensajeBoletas("")
Set oGen = Nothing
Print #nFicSal, PrnSet("B+") + ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) & ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) + PrnSet("B-")
Print #nFicSal, ""
Print #nFicSal, ""
Print #nFicSal, ""

Close nFicSal
Exit Sub
ERROR:
    Close nFicSal
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function ImprimeArqueo(ByVal psNomCmact As String, ByVal pdFecha As Date, _
                             ByVal psMovNro As String, ByVal psOpecod As String, _
                             ByVal rsBillIngreso As ADODB.Recordset, ByVal rsMonIng As ADODB.Recordset, _
                             ByVal rsTotales As ADODB.Recordset, ByVal rsOP As ADODB.Recordset, _
                             ByVal psAreaAgeCh As String, ByVal psCajaChicaNomb As String, _
                             ByVal psCodPersResp As String, ByVal psResponsable As String, _
                             ByVal psAgeCod As String, ByVal psAgencia As String, _
                             ByVal psCodPers As String, ByVal psPersNombre As String, _
                             ByVal psHoraIni As String, ByVal psHoraFin As String, _
                             ByVal pnNumPag As Integer, ByVal pnNumCol As Integer) As String
            

Dim lsTexto As String
Dim lnTotalBill As Currency
Dim lnTotalMon As Currency
Dim lsCab1 As String
Dim lsCab2 As String
Dim lsTitulo2 As String
Dim lnNumPag As Integer
On Error GoTo ImprimeArqueoErr

lnTotalBill = 0
lnTotalMon = 0
lsTexto = PrnSet("MI", 10)
lsCab1 = CON + ImpreFormat("Resp. Caja   : [" & psCodPersResp & "] " + psResponsable, (pnNumCol / 2) + 65) + BON + ImpreFormat("Inicio  :" & psHoraIni, 25) + COFF + BOFF
lsCab2 = CON + ImpreFormat("Resp. Arqueo : [" & psCodPers & "] " + psPersNombre, (pnNumCol / 2) + 65) + BON + ImpreFormat("Término :" & psHoraFin, 25) + COFF + BOFF
lsTitulo2 = PrnSet("I+") + "Caja Chica N° [" & psAreaAgeCh & "] " & psCajaChicaNomb + PrnSet("I-")
lnNumPag = 0
'''Linea lsTexto, CabeRepo(psNomCmact, psAgencia, "CONTABILIDAD", IIf(Mid(psOpeCod, 3, 1) = "1", "SOLES", "DOLARES"), Format(pdFecha, gsFormatoFechaView), "A R Q U E O   D E   C A J A   C H I C A", lsTitulo2, "", "", lnNumPag, pnNumCol) 'MARG ERS044-2016
Linea lsTexto, CabeRepo(psNomCmact, psAgencia, "CONTABILIDAD", IIf(Mid(psOpecod, 3, 1) = "1", StrConv(gcPEN_PLURAL, vbUpperCase), "DOLARES"), Format(pdFecha, gsFormatoFechaView), "A R Q U E O   D E   C A J A   C H I C A", lsTitulo2, "", "", lnNumPag, pnNumCol) 'MARG ERS044-2016
Linea lsTexto, lsCab1
Linea lsTexto, lsCab2
Linea lsTexto, ""
Linea lsTexto, BON + ImpreFormat("FONDOS RECONTADOS ", pnNumCol)
Linea lsTexto, ImpreFormat("I. BILLETES ", pnNumCol)
Linea lsTexto, String(pnNumCol - 20, "-")
Linea lsTexto, ImpreFormat("VALOR", 27) + ImpreFormat("CANTIDAD", 12, 0, True) & ImpreFormat("IMPORTE", 15, 2, True)
Linea lsTexto, String(pnNumCol - 20, "-") + BOFF
If Not rsBillIngreso Is Nothing Then
    If rsBillIngreso.State = adStateOpen Then
        If Not rsBillIngreso.EOF And Not rsBillIngreso.BOF Then
            Do While Not rsBillIngreso.EOF
                Linea lsTexto, ImpreFormat(rsBillIngreso(0), 20) + ImpreFormat(rsBillIngreso(1), 10, 0, True) & ImpreFormat(rsBillIngreso(2), 15, 2, True)
                lnTotalBill = lnTotalBill + rsBillIngreso(2)
                rsBillIngreso.MoveNext
            Loop
        End If
        'rsBillIngreso.Close: Set rsBillIngreso = Nothing
    End If
End If
Linea lsTexto, BON + String(pnNumCol - 20, "-")
Linea lsTexto, ImpreFormat("TOTAL :", 20) + ImpreFormat("", 10, 0, True) & ImpreFormat(lnTotalBill, 15, 2, True) + BOFF

Linea lsTexto, ""
Linea lsTexto, BON + ImpreFormat("II. MONEDAS ", pnNumCol)
Linea lsTexto, String(pnNumCol - 20, "-")
Linea lsTexto, ImpreFormat("VALOR", 27) + ImpreFormat("CANTIDAD", 12, 0, True) & ImpreFormat("IMPORTE", 15, 2, True)
Linea lsTexto, String(pnNumCol - 20, "-") + BOFF

If Not rsMonIng Is Nothing Then
    If rsMonIng.State = adStateOpen Then
        If Not rsMonIng.EOF And Not rsMonIng.BOF Then
            Do While Not rsMonIng.EOF
                Linea lsTexto, ImpreFormat(rsMonIng(0), 20) + ImpreFormat(rsMonIng(1), 10, 0, True) & ImpreFormat(rsMonIng(2), 15, 2, True)
                lnTotalMon = lnTotalMon + rsMonIng(2)
                rsMonIng.MoveNext
            Loop
        End If
        'rsMonIng.Close: Set rsMonIng = Nothing
    End If
End If
Linea lsTexto, BON + String(pnNumCol - 20, "-")
Linea lsTexto, ImpreFormat("TOTAL :", 20) + ImpreFormat("", 10, 0, True) & ImpreFormat(lnTotalMon, 15, 2, True) + BOFF
Linea lsTexto, ""
If Not rsOP Is Nothing Then
    If rsOP.State = adStateOpen Then
        If Not rsOP.EOF And Not rsOP.BOF Then
        Linea lsTexto, BON + ImpreFormat("III. ORDENES DE PAGO PENDIENTE DE COBRO", pnNumCol)
        Linea lsTexto, String(pnNumCol - 20, "-")
        Linea lsTexto, ImpreFormat("N°", 10) + ImpreFormat("ORDEN", 20) & ImpreFormat("IMPORTE", 15, 2, True)
        Linea lsTexto, String(pnNumCol - 20, "-") + BOFF
            Do While Not rsOP.EOF
                Linea lsTexto, ImpreFormat(rsOP.Bookmark, 10) + ImpreFormat(rsOP!Orden, 20) & ImpreFormat(rsOP!Importe, 15, 2, True)
                rsOP.MoveNext
            Loop
        End If
        rsOP.Close: Set rsOP = Nothing
    End If
End If
If Not rsTotales Is Nothing Then
    If rsTotales.State = adStateOpen Then
        If Not rsTotales.EOF And Not rsTotales.BOF Then
            Linea lsTexto, "", 1
            Do While Not rsTotales.EOF
                If rsTotales.Bookmark = 7 Or rsTotales.Bookmark = 9 Then
                    lsTexto = lsTexto + BON
                End If
                Linea lsTexto, ImpreFormat(rsTotales(0), 40) + ImpreFormat(rsTotales(1), 15, 2, True) + BOFF
                rsTotales.MoveNext
            Loop
        End If
    End If
End If
Linea lsTexto, "", 2
Linea lsTexto, "Siendo las " & psHoraFin & " se dio por terminado el arqueo de Caja,"
Linea lsTexto, "devolviendose al responsable " & psResponsable & " todo el dinero y documentos de Caja,"
Linea lsTexto, "tal como se nos entregó para efecto del arqueo."
Linea lsTexto, "Se firma la presente en señal de conformidad"
'Linea lsTexto, "", 2
Linea lsTexto, ImprePiePag(pnNumCol, "69")

ImprimeArqueo = lsTexto

Exit Function
ImprimeArqueoErr:
    Err.Raise vbObjectError + 100, "ImprimeArqueo", Err.Description & " Error ImprimeArqueo"
End Function

Public Function ImprimeBoleta(ByVal psTit As String, ByVal psText As String, ByVal psCodOpe As String, ByVal pnMonto As String, _
            ByVal psCliente As String, ByVal psCodCta As String, ByVal psNumDoc As String, ByVal pnSaldo As Double, _
            ByVal pnInteresA As String, NomDoc As String, ByVal pnNumExt As Long, ByVal pnSaldoC As Double, _
            Optional pbOpSaldoC As Boolean = True, Optional pbSaldoInt As Boolean = True, Optional psNumDias As String = "----", _
            Optional psNomAgeRem As String = "", Optional psCodUsuRem As String = "", Optional pbCuenta As Boolean = False, _
            Optional psComCmac As String = "XXX", Optional psLin3 As String = "XXX", Optional psTexto As String = "XXX", _
            Optional pdFecsis As Date, Optional psNomAge As String = "", Optional psCodUser As String) As String

Dim nFicSal As Integer
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMonto As String
Dim sSDisp As String
Dim sIntAcum As String
Dim sMax As Integer
Dim sAux As Integer
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

ETIQ:

On Error GoTo ERROR

lnTope = 0 '6 'Tope de lineas en Boleta

lsNegritaOn = oImpresora.gPrnBoldON
lsNegritaOff = oImpresora.gPrnBoldOFF

lsNroExt = Str(pnNumExt)

nFicSal = FreeFile
Open sLPT For Output As nFicSal

Print #nFicSal, oImpresora.gPrnInicializa;

Print #nFicSal, oImpresora.gPrnEspaLineaN;   'espaciamiento lineas 1/6 pulg.
Print #nFicSal, oImpresora.gPrnTamPagina22;   'Longitud de página a 22 líneas'
Print #nFicSal, oImpresora.gPrnTamLetra10CPI;   'Tamaño 10 cpi
Print #nFicSal, oImpresora.gPrnTpoLetraRoman;      'Tipo de Letra Sans Serif
Print #nFicSal, oImpresora.gPrnCondensadaOFF  ' cancela condensada
Print #nFicSal, oImpresora.gPrnBoldOFF  ' desactiva negrita

sSep = 15
sIni = 1
sMax = 33
sAux = 5


sFecha = Format$(pdFecsis, gsFormatoFechaView)
sHora = Format$(Time, "hh:mm:ss")
sMonto = Format$(pnMonto, "#,##0.00")
sSDisp = Format$(pnSaldo, "#,##0.00")
lsSaldoC = Format$(pnSaldoC, "#,##0.00")

lsNomAge = psNomAge

'Print #nFicSal, oImpresora.gPrnSaltoLinea;
Print #nFicSal, lsNegritaOn; 'Activa Negrita
Print #nFicSal, Tab(sIni); "CMACT - AHORRO"; Space(19 + sSep + sAux); "CMACT - AHORRO"

If Mid(psCodCta, 9, 1) = 1 Then
    '''Print #nFicSal, Tab(sIni); Trim(psNomAge) & "-SOLES"; Space(sAux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; Space(sSep); Trim(psNomAge) & "-SOLES"; Space(sAux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; 'MARG ERS044-2016
    Print #nFicSal, Tab(sIni); Trim(psNomAge) & "-" & StrConv(gcPEN_PLURAL, vbUpperCase); Space(sAux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN_PLURAL, vbUpperCase))) + lsNroExt; Space(sSep); Trim(psNomAge) & "-" & StrConv(gcPEN_PLURAL, vbUpperCase); Space(sAux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN_PLURAL, vbUpperCase))) + lsNroExt; 'MARG ERS044-2016
Else
    Print #nFicSal, Tab(sIni); Trim(psNomAge) & "-DOLARES"; Space(sAux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt; Space(sSep); Trim(psNomAge) & "-DOLARES"; Space(sAux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt;
End If

If psNomAgeRem = "" Then
    Print #nFicSal, ""
Else
    Print #nFicSal, Tab(sIni); "Ag.Rem: " & Trim(psNomAgeRem); Space(sAux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))); "Ag.Rem: " & Trim(psNomAgeRem)
End If

If psComCmac = "XXX" Then
    If psLin3 = "XXX" Then
        Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
    Else
        Print #nFicSal, Tab(sIni); psLin3 & Space(sAux + sSep + sMax - Len(psLin3)) & psLin3 & lsNegritaOff;  'Desactiva Negrita
        lnNumLinCmac = 1
    End If
    lnNumLinCmac = 0
Else
    Print #nFicSal, Tab(sIni); "NroDocCmac:" & psComCmac & Space(sAux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & "NroDocCmac:" & psComCmac & lsNegritaOff;  'Desactiva Negrita
    lnNumLinCmac = 1
End If

Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(sAux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora

'psCliente = PstaNombre(psCliente)

lnCliAux = InStr(1, psCliente, "*", vbTextCompare)

If lnCliAux = 0 Then
    If sAux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + sAux)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(psCliente); Space(sAux + sMax + sSep - Len(psCliente)); ImpreCarEsp(psCliente)
Else
    lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
    lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
    
    If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + sAux)
    If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + sAux)
    
    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux1); Space(sAux + sMax + sSep - Len(lsCliAux1)); ImpreCarEsp(lsCliAux1)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux2); Space(sAux + sMax + sSep - Len(lsCliAux2)); ImpreCarEsp(lsCliAux2)
    
    lnCliAux = 1
End If

If pbSaldoInt Or pbCuenta Then
    Print #nFicSal, Tab(sIni); "Cuenta:" & psCodCta; Space(8 + sSep + sAux); "Cuenta:" & psCodCta
Else
    Print #nFicSal, ""
End If

psTit = Trim(psTit)
psTit = CentrarCadena(psTit, 28)
Print #nFicSal, lsNegritaOn; 'Activa Negrita
Print #nFicSal, Tab(sIni + 1); "-----" & psTit & "-----"; Space(-1 + sSep); "-----" & psTit & "-----"


lnChq = InStr(1, psText, "*", vbTextCompare)

If psTexto = "XXX" Then
    If lnChq = 0 Then
        Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(Mid(psText, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)); sMonto; Space(-1 + sSep); ImpreCarEsp(Trim(Mid(psText, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)); sMonto
        Print #nFicSal, ""
    Else
        lsChqAux1 = (Mid(psText, 1, lnChq - 1))
        lsChqAux2 = (Mid(psText, lnChq + 1))
        Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(lsChqAux1)); Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)); sMonto; Space(-1 + sSep); ImpreCarEsp(Trim(lsChqAux1)); Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)); sMonto
        Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(lsChqAux2)); Space(sMax + 6 - Len(Trim(lsChqAux2))); Space(-1 + sSep); ImpreCarEsp(Trim(lsChqAux2)); Space(sMax + 6 - Len(Trim(lsChqAux2)))
    End If
Else
    Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(psTexto)); Space(sAux + sSep + sMax - Len(Trim(psTexto))); ImpreCarEsp(Trim(psTexto))
    Print #nFicSal, ""
End If


Print #nFicSal, lsNegritaOff; 'Desactiva Negrita

If pbSaldoInt Then
    If MsgBox("Desea Imprimir el Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Print #nFicSal, Tab(sIni); "Saldo Disponible"; Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)); sSDisp; Space(-1 + sSep); "Saldo Disponible"; Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)); sSDisp
        If pbOpSaldoC Then
            Print #nFicSal, Tab(sIni); "Saldo Contable"; Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)); lsSaldoC; Space(-1 + sSep); "Saldo Contable"; Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)); lsSaldoC
        Else
            Print #nFicSal, ""
        End If
    Else
        Print #nFicSal, ""
        Print #nFicSal, ""
        pbSaldoInt = False
    End If
Else
    Print #nFicSal, ""
End If

lsInteres = pnInteresA

If pbSaldoInt Then
    If lsInteres <> "No Valido" Then
        lsInteres = Format(lsInteres, "#,##0.00")
        Print #nFicSal, Tab(sIni); "Interes del Mes"; Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)); lsInteres; Space(-1 + sSep); "Interes del Mes"; Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)); lsInteres
    End If
Else
    Print #nFicSal, ""
End If

If Not psNumDoc = "" Then
    Print #nFicSal, Tab(sIni); NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc; Space(-1 + sSep); NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc
Else
    Print #nFicSal, ""
End If

If Not psNumDias = "----" Then
    Print #nFicSal, Tab(sIni); "Nro Dias Interes"; Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)); psNumDias; Space(-1 + sSep); "Nro Dias Interes"; Space(sMax + 6 - Len(psNumDias) - Len("Nro Dias Interes")); psNumDias
    lnTope = 4 - lnCliAux
Else
    lnTope = 3 - lnCliAux
End If

Print #nFicSal, Tab(sIni); "---------------------------------------"; Space(-1 + sSep); "---------------------------------------"
If psCodUsuRem = "" Then
    Print #nFicSal, Tab(sIni); ImpreCarEsp(psCodUser); Space(29 + sSep + sAux); ImpreCarEsp(psCodUser)
Else
    Print #nFicSal, Tab(sIni); "Loc/Rem"; Space(sMax + sAux - Len("Loc/Rem") - 1 - 8); ImpreCarEsp(psCodUser) & "/"; ImpreCarEsp(psCodUsuRem); Space(sSep); "Loc/Rem"; Space(sMax + sAux - Len("Loc/Rem") - 1 - 8); ImpreCarEsp(psCodUser) & "/"; ImpreCarEsp(psCodUsuRem)
End If
Dim clsGen As DGeneral
Set clsGen = New DGeneral
lsMensaje = clsGen.GetMensajeBoletas(psCodCta)
Set clsGen = Nothing
Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsMensaje); Space(-1 + sSep); ImpreCarEsp(lsMensaje); lsNegritaOff

lnNumLinCmac = lnNumLinCmac + 1
 
'Print #nFicSal, oImpresora.gPrnSaltoPagina
For sAux = 1 To (lnTope - lnNumLinCmac)
    Print #nFicSal, ""
Next sAux
Close nFicSal
Exit Function
ERROR:
    Close nFicSal
    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
        GoTo ETIQ
    End If
End Function


'Public Function ImprimeLibroDiario(ByVal dFecIni As Date, ByVal dFecFin As Date, psAge As String, psOpe As String, pnLinPage As Integer) As String
'Dim sSQL As String
'Dim R As New ADODB.Recordset
'Dim P As Integer, nLi As Integer
'Dim sCadRep As String
'Dim sCad    As String
'Dim sClave  As String
'Dim nTotalDocD As Currency
'Dim nTotalDocH As Currency
'Dim nTotalD    As Currency
'Dim nTotalH    As Currency
'Dim sGlosa     As String
'Dim sDesOp     As String
'
'Dim sMov As String
'Dim sFec As String
'Dim sAge As String
'Dim sTpo As String
'
'    On Error GoTo ImprimeLibroDiarioErr
'    nTotalDocD = 0
'    nTotalDocH = 0
'    nTotalD = 0: nTotalH = 0
'    sCadRep = ""
'    P = 0
'    Dim oFun As New NContFunciones
'    Set R = oFun.GetLibroDiario(dFecIni, dFecFin, psAge, psOpe)
'    Set oFun = Nothing
'    If R.EOF Then
'       Err.Raise 50001, "NContImprimir", "No se registraron Asientos para generar Libro Diario"
'       R.Close: Set R = Nothing
'       Exit Function
'    End If
'      sCadRep = PrnSet("MI", 8)
'      nLi = pnLinPage
'      RaiseEvent BarraShow(R.RecordCount)
'      sClave = ""
'      sDesOp = R!cOpeGruDesc
'      sGlosa = R!cMovDesc
'      sMov = R!cMovNro
'      Do While Not R.EOF
'         DoEvents
'         RaiseEvent BarraProgress(R.Bookmark, "LIBRO DIARIO", "", "Procesando... ", vbBlue)
'         sFec = Mid(R!cMovNro, 7, 2) & "/" & Mid(R!cMovNro, 5, 2) & "/" & Mid(R!cMovNro, 1, 4)
'         sAge = Mid(R!cMovNro, 18, 2)         'Agencia
'         sTpo = R!cOpeGruCod                  'Tpo de Asiento
'         Linea sCadRep, ImprimeLibroDiarioCab(nLi, P, dFecIni, dFecFin, R!nMovCorrela, sFec, sAge, sTpo, sMov, pnLinPage), 0
'         If sMov <> R!cMovNro Then
'            Linea sCadRep, BON & Space(5) + "Total de Documentos: " + Space(60) & PrnVal(nTotalDocD, 16, 2) + "  " & PrnVal(nTotalDocH, 16, 2) & BOFF, , nLi
'            Linea sCadRep, BON & Space(5) + "Glosa : " + ImpreGlosa(sGlosa, 135, "Glosa : ", , False, nLi) & BOFF, 2, nLi
'            Linea sCadRep, BON & "Documento: " & Format(R!nMovCorrela, "#####0") & "    Tipo de Asiento: " & R!cOpeGruCod & "     Agencia: " & sAge & "     Fecha: " & sFec & "    Mov.Nro: " & R!cMovNro & BOFF, 2, nLi
'            sGlosa = R!cMovDesc
'            sMov = R!cMovNro
'
'            nTotalDocH = 0
'            nTotalDocD = 0
'         End If
'         Linea sCadRep, ImprimeLibroDiarioCab(nLi, P, dFecIni, dFecFin, R!nMovCorrela, sFec, sAge, sTpo, sMov, pnLinPage), 0
'
'         ' Linea de Detalle
'         If R!nMovImporte > 0 Then
'            Linea sCadRep, Space(5) + Justifica(R!cCtaContCod, 22) & " " & Justifica(R!cCtaContDesc, 60) & PrnVal(R!nMovImporte, 16, 2)
'            nTotalDocD = nTotalDocD + R!nMovImporte
'            nTotalD = nTotalD + R!nMovImporte
'         Else
'            Linea sCadRep, Space(5) + Justifica(R!cCtaContCod, 22) & " " & Justifica(R!cCtaContDesc, 60) & Space(18) & PrnVal(R!nMovImporte * -1, 16, 2)
'            nTotalDocH = nTotalDocH + (R!nMovImporte * -1)
'            nTotalH = nTotalH + (R!nMovImporte * -1)
'         End If
'         nLi = nLi + 1
'         If P Mod 20 = 0 Then
'            sCad = sCad & sCadRep
'            sCadRep = ""
'         End If
'         R.MoveNext
'      Loop
'      R.MoveLast
'      Linea sCadRep, ImprimeLibroDiarioCab(nLi, P, dFecIni, dFecFin, R!nMovCorrela, sFec, sAge, sTpo, sMov, pnLinPage), 0
'      Linea sCadRep, BON & Space(5) + "Total de Documentos: " + Space(60) & PrnVal(nTotalDocD, 16, 2) + "  " & PrnVal(nTotalDocH, 16, 2) & BOFF
'      Linea sCadRep, BON & Space(5) + "Glosa : " + sGlosa & BOFF, 3
'      Linea sCadRep, BON & "     T O T A L E S  " & Space(66) & PrnVal(nTotalD, 16, 2) & "  " & PrnVal(nTotalH, 16, 2) & BOFF
'      sCad = sCad & sCadRep
'      R.Close: Set R = Nothing
'    ImprimeLibroDiario = sCad
'    RaiseEvent BarraClose
'
'    Exit Function
'ImprimeLibroDiarioErr:
'    RaiseEvent BarraClose
'    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeLibroDiario Method")
'End Function

'Private Function ImprimeLibroDiarioCab(nLi As Integer, P As Integer, dFecIni As Date, dFecFin As Date, nNro As Currency, sFec As String, sAge As String, sTpo As String, sMov As String, pnLinPage As Integer) As String
'Dim sCadCab As String
'If pnLinPage - 6 < nLi Then
'   Linea sCadCab, COFF & BON & Cabecera("L I B R O   D I A R I O ", P, "", 80, "", Format(dFecFin, gsFormatoFechaView)), 2
'   Linea sCadCab, CON & Space(5) + "CODIGO" + Space(15) + "NOMBRE" + Space(64) + "DEBE" + Space(13) + "HABER", 2
'   Linea sCadCab, "Documento: " & Format(nNro, "#####0") & "    Tipo de Asiento: " & sTpo & "     Agencia: " & sAge & "     Fecha: " & sFec & "    Mov.Nro: " & sMov & BOFF, 2
'   nLi = 9
'End If
'ImprimeLibroDiarioCab = sCadCab
'End Function

Public Function ImprimeLibroDiario(ByVal dFecIni As Date, ByVal dFecFin As Date, psAge As String, psOpe As String, pnLinPage As Integer, Optional pbHistoCtaContb As Boolean = False) As String 'JACA 20111229 se agregó pbHistoCtaContb
Dim sSQL As String
Dim R As New ADODB.Recordset
Dim P As Integer, nLi As Integer
Dim sCadRep As String
Dim sCad    As String
Dim sClave  As String
Dim nTotalDocD As Currency
Dim nTotalDocH As Currency
Dim nTotalD    As Currency
Dim nTotalH    As Currency
Dim sGlosa     As String

Dim sFec As String
Dim sAge As String
Dim sTpo As String
Dim sMovi As String


    On Error GoTo ImprimeLibroDiarioErr
    nTotalDocD = 0
    nTotalDocH = 0
    nTotalD = 0: nTotalH = 0
    sCadRep = ""
    P = 0
    Dim oFun As New NContFunciones
    RaiseEvent BarraShow(1)
    RaiseEvent BarraProgress(0, "LIBRO DIARIO", "", "Cargando Datos... ", vbBlue)
    'Set R = oFun.GetLibroDiario(dFecIni, dFecFin, psAge, psOpe, pbHistoCtaContb) 'JACA 20111229 se agregó pbHistoCtaContb
    Set R = oFun.GetLibroDiario(dFecIni, dFecFin, psAge, psOpe) 'JACA 20111229 se agregó pbHistoCtaContb
    Set oFun = Nothing
    RaiseEvent BarraClose
    
    If R.State = 0 Then
        Exit Function
    End If
    
    If R.EOF Then
       Err.Raise 50001, "NContImprimir", "No se registraron Asientos para generar Libro Diario"
       R.Close: Set R = Nothing
       Exit Function
    End If
      'by Capi 07102008
      'sCadRep = PrnSet("MI", 8)
      nLi = pnLinPage
      RaiseEvent BarraShow(R.RecordCount)
      sClave = ""
      sGlosa = Justifica(R!cOpeGruDesc, 20) + Space(2) + "Descripcion: " + Justifica(R!cMovdesc, 82)

      Do While Not R.EOF
         DoEvents
         RaiseEvent BarraProgress(R.Bookmark, "LIBRO DIARIO", "", "Procesando... ", vbBlue)
         sFec = Mid(R!cMovNro, 7, 2) & "/" & Mid(R!cMovNro, 5, 2) & "/" & Mid(R!cMovNro, 1, 4)
         sAge = Mid(R!cMovNro, 18, 2)         'Agencia
         sTpo = R!cOpeGruCod                  'Tpo de Asiento
         sMovi = R!cMovNro
         
         If sClave <> sFec & sAge & sTpo & sMovi Then
            Linea sCadRep, ImprimeLibroDiarioCab(nLi, P, dFecIni, dFecFin, R!nMovCorrela, sFec, sAge, sTpo, pnLinPage, sClave), 0
            If sClave <> "" Then
               'By Capi 01102008
               
'               Linea sCadRep, BON & Space(5) + "Total de Documentos: " + Space(37) & PrnVal(nTotalDocD, 14, 2) + " " & PrnVal(nTotalDocH, 14, 2) & BOFF
'               Linea sCadRep, BON & Space(5) + "Glosa : " + sGlosa & BOFF, 3
'               Linea sCadRep, BON & "Documento: " & Format(R!nMovCorrela, "#####0") & "     Tipo de Asiento: " & R!cOpeGruCod & "      Agencia: " & sAge & "      Fecha: " & sFec & BOFF, 2
'               sGlosa = Justifica(R!cOpeGruDesc, 20) + Space(2) + "Descripcion: " + Justifica(R!cMovdesc, 82)
'               ALPA 20091118*******************************************
'               Linea sCadRep, CON & Space(5) + "Total de Documentos: " + Space(37) & PrnVal(nTotalDocD, 14, 2) + " " & PrnVal(nTotalDocH, 14, 2)
'               Linea sCadRep, CON & Space(5) + "Glosa : " + Left(sGlosa, 100), 3
'               Linea sCadRep, CON & "Documento: " & Format(R!nMovCorrela, "#####0") & "     Tipo de Asiento: " & R!cOpeGruCod & "      Agencia: " & sAge & "      Fecha: " & sFec, 2
               Linea sCadRep, Space(5) + "Total de Documentos: " + Space(37) & PrnVal(nTotalDocD, 14, 2) + " " & PrnVal(nTotalDocH, 14, 2)
               Linea sCadRep, Space(5) + "Glosa : " + Left(sGlosa, 100), 3
'
               'ALPA 20120410******************************
               'Linea sCadRep, "Documento: " & Format(R!nMovCorrela, "#####0") & "     Tipo de Asiento: " & R!cOpeGruCod & "      Agencia: " & sAge & "      Fecha: " & sFec & " Linea: " & "(" & CStr(nLi) & ")", 2
               Linea sCadRep, "Documento: " & Format(R!nMovCorrela, "#####0") & "     Tipo de Asiento: " & R!cOpeGruCod & "      Agencia: " & sAge & "      Fecha: " & sFec, 2
               '*******************************************
'
               '********************************************************
               sGlosa = Justifica(R!cOpeGruDesc, 20) + Space(2) + "Descripcion: " + Justifica(R!cMovdesc, 82)
               nLi = nLi + 7
               nTotalDocH = 0
               nTotalDocD = 0
               '*********************************************
            End If
            sClave = sFec & sAge & sTpo & sMovi
         End If
         Linea sCadRep, ImprimeLibroDiarioCab(nLi, P, dFecIni, dFecFin, R!nMovCorrela, sFec, sAge, sTpo, pnLinPage), 0
            
         ' Linea de Detalle
         If R!nMovImporte > 0 Then
            Linea sCadRep, Space(2) + Justifica(R!cCtaContCod, 20) & " " & Justifica(R!cCtaContDesc, 40) & PrnVal(R!nMovImporte, 14, 2) & Space(17)
            nTotalDocD = nTotalDocD + R!nMovImporte
            nTotalD = nTotalD + R!nMovImporte
         Else
            Linea sCadRep, Space(2) + Justifica(R!cCtaContCod, 20) & " " & Justifica(R!cCtaContDesc, 40) & Space(15) & PrnVal(R!nMovImporte * -1, 14, 2)
            nTotalDocH = nTotalDocH + (R!nMovImporte * -1)
            nTotalH = nTotalH + (R!nMovImporte * -1)
         End If
         nLi = nLi + 1
         If P Mod 20 = 0 Then
            sCad = sCad & sCadRep
            sCadRep = ""
         End If
         R.MoveNext
      Loop
      R.MoveLast
      'By Capi 07082008
'      Linea sCadRep, ImprimeLibroDiarioCab(nLi, P, dFecIni, dFecFin, R!nMovCorrela, sFec, sAge, sTpo, pnLinPage), 0
'      Linea sCadRep, BON & Space(5) + "Total de Documentos: " + Space(37) & PrnVal(nTotalDocD, 14, 2) + " " & PrnVal(nTotalDocH, 14, 2) & BOFF
'      Linea sCadRep, BON & Space(5) + "Glosa : " + sGlosa & BOFF, 3
'      Linea sCadRep, BON & "     T O T A L E S  " & Space(43) & PrnVal(nTotalD, 14, 2) & " " & PrnVal(nTotalH, 14, 2) & BOFF
'     ALPA 20091118********************************************
      Linea sCadRep, ImprimeLibroDiarioCab(nLi, P, dFecIni, dFecFin, R!nMovCorrela, sFec, sAge, sTpo, pnLinPage), 0
      Linea sCadRep, Space(5) + "Total de Documentos: " + Space(37) & PrnVal(nTotalDocD, 14, 2) + " " & PrnVal(nTotalDocH, 14, 2)
      Linea sCadRep, Space(5) + "Glosa : " + Left(sGlosa, 100), 3
      Linea sCadRep, "     T O T A L E S  " & Space(43) & PrnVal(nTotalD, 14, 2) & " " & PrnVal(nTotalH, 14, 2)
'      Linea sCadRep, CON & ImprimeLibroDiarioCab(nLi, P, dFecIni, dFecFin, R!nMovCorrela, sFec, sAge, sTpo, pnLinPage), 0
'      Linea sCadRep, CON & Space(5) + "Total de Documentos: " + Space(37) & PrnVal(nTotalDocD, 14, 2) + " " & PrnVal(nTotalDocH, 14, 2)
'      Linea sCadRep, CON & Space(5) + "Glosa : " + Left(sGlosa, 100), 3
'      Linea sCadRep, CON & "     T O T A L E S  " & Space(43) & PrnVal(nTotalD, 14, 2) & " " & PrnVal(nTotalH, 14, 2)
      '********************************************************
      '
      sCad = sCad & sCadRep
      R.Close: Set R = Nothing
    ImprimeLibroDiario = sCad
    RaiseEvent BarraClose
    
    Exit Function
ImprimeLibroDiarioErr:
    RaiseEvent BarraClose
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeLibroDiario Method")
End Function
'By capi 23112008
'Private Function ImprimeLibroDiarioCab(nLi As Integer, P As Integer, dFecIni As Date, dFecFin As Date, nNro As Currency, sFec As String, sAge As String, sTpo As String, pnLinPage As Integer) As String
Private Function ImprimeLibroDiarioCab(nLi As Integer, P As Integer, dFecIni As Date, dFecFin As Date, nNro As Currency, sFec As String, sAge As String, sTpo As String, pnLinPage As Integer, Optional ByVal psClave As String) As String

Dim sCadCab As String
If pnLinPage - 6 < nLi Then
   'By Capi 07102008
   'Linea sCadCab, COFF & BON & Cabecera("L I B R O   D I A R I O ", P, "", 80, "", Format(dFecFin, gsFormatoFechaView), gsNomCmac), 2
   'Linea sCadCab, CON & Space(5) + "CODIGO" + Space(15) + "NOMBRE" + Space(38) + "DEBE" + Space(8) + "HABER", 2
   'Linea sCadCab, "Documento: " & Format(nNro, "#####0") & "     Tipo de Asiento: " & sTpo & "      Agencia: " & sAge & "      Fecha: " & sFec & BOFF, 2
   'ALPA 20091118**************************************
   Linea sCadCab, Cabecera("L I B R O   D I A R I O ", P, "", 80, "", Format(dFecFin, gsFormatoFechaView), gsNomCmac), 2
   Linea sCadCab, Space(5) + "CODIGO" + Space(15) + "NOMBRE" + Space(38) + "DEBE" + Space(8) + "HABER", 2
'   Linea sCadCab, CON & Cabecera("L I B R O   D I A R I O ", P, "", 80, "", Format(dFecFin, gsFormatoFechaView), gsNomCmac), 2
'   Linea sCadCab, Space(5) + "CODIGO" + Space(15) + "NOMBRE" + Space(38) + "DEBE" + Space(8) + "HABER", 2
   '***************************************************
   If psClave = "" Then
        Linea sCadCab, "Documento: " & Format(nNro, "#####0") & "     Tipo de Asiento: " & sTpo & "      Agencia: " & sAge & "      Fecha: " & sFec, 2
   End If
   nLi = 9
End If
ImprimeLibroDiarioCab = sCadCab
End Function

Public Function ImprimeLibroMayor(dFechaDel As Date, dFechaAl As Date, psAge As String, psOpe As String, pnLinPage As Integer, Optional ByVal pnResDet As Integer = 0, Optional pbHistoCtaContb As Boolean = False) As String 'JACA 20111229 se agregó pbHistoCtaContb
Dim rsCta As New ADODB.Recordset
Dim nImporte As Currency
Dim nSaldo As Currency, nSaldoIni As Currency
Dim nDebe  As Currency, nHaber  As Currency
Dim nDebeD As Currency, nHaberH As Currency
Dim nDebeT As Currency, nHaberT As Currency
Dim nHaberD As Currency
Dim sTexto As String, sImpre As String
Dim sFecha As String
Dim sSQL   As String
Dim nItem  As Integer
Dim nLin   As Integer, P As Integer
Dim sCtaCod As String, sCtaDes As String
Dim sTipoCta As String
Dim oConec As New DConecta
Dim oSdo   As NCtasaldo
    On Error GoTo ImprimeLibroMayorErr

nItem = 0: nLin = pnLinPage: P = 0
If dFechaDel > dFechaAl Then
   Err.Raise 50001, "ImprimeLibroMayor", "Fecha Inicial debe ser menor o igual que fecha final."
   Exit Function
End If

RaiseEvent BarraShow(1)
RaiseEvent BarraProgress(0, "LIBRO MAYOR", "", "Cargando datos... ", vbBlue)

'*** PEAC 20110406
'sSQL = "SELECT a.cCtaContCod, c.cCtaContDesc, IsNull(m1.nMovCorrela,0) nMovCorrela, b.cMovNro, OT.cOpeGruCod, cc.cCtaCaracter, ISNULL(sdo.nCtaSaldoImporte,0) nSaldoIni, Left(b.cMovDesc,82) As cMovDesc," _
     & "       ISNULL(CASE WHEN a.nMovImporte > 0 THEN a.nMovImporte END,0) as nDebe, " _
     & "       ISNULL(CASE WHEN a.nMovImporte < 0 THEN a.nMovImporte * -1 END,0) as nHaber " _
     & "FROM   Mov b    JOIN MovCta a ON a.nMovNro = b.nMovNro JOIN " & vsBaseComunes & "CtaCont c ON c.cCtaContCod = a.cCtaContCod" _
     & "                Left JOIN MovCont m1 ON m1.nMovNro = b.nMovNro " _
     & "                JOIN " & vsBaseComunes & "OpeTpo OT ON OT.cOpeCod = b.cOpeCod " _
     & "                JOIN " & vsBaseComunes & "CtaContClase cc ON c.cCtaContCod LIKE RTRIM(cc.cCtaContCod) + '%' " _
     & "           LEFT JOIN (SELECT cCtaContCod, nCtaSaldoImporte FROM CtaSaldo cs WHERE dCtaSaldoFecha = " _
     & "                             (SELECT MAX(dCtaSaldoFecha) FROM CtaSaldo cs1 WHERE cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha < '" & Format(dFechaDel, gsFormatoFecha) & "') " _
     & "                     ) Sdo ON Sdo.cCtaContCod = a.cCtaContCod " _
     & "WHERE  b.nMovEstado = " & gMovEstContabMovContable & " and not b.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & ") and Not b.cOpecod Like '7018%' and " _
     & "       substring(b.cmovnro,1,8) BETWEEN '" & Format(dFechaDel, "yyyymmdd") & "' and '" & Format(dFechaAl, "yyyymmdd") & "' " _
     & "       " & IIf(psAge = "", "", " And substring(b.cmovnro,18,2) = '" & psAge & "'  ") & " " _
     & "ORDER BY a.cCtaContCod, m1.nMovCorrela, SubString(b.cMovNro,1,8), Substring(b.cMovNro,18,2), OT.cOpeGruCod, b.cMovNro "

'JACA 20111229 se modificó para agregar ctas contable historicas
If Not pbHistoCtaContb Then
    sSQL = "exec stp_sel_ImprimeLibroMayor '" & Format(dFechaDel, gsFormatoFecha) & "','" & Format(dFechaDel, "yyyymmdd") & "','" & Format(dFechaAl, "yyyymmdd") & "','" & IIf(psAge = "", "%%", psAge) & "'," & IIf(pnResDet = 2, 1, pnResDet) 'PASI TI-ERS086-2014
Else
    'sSQL = "exec Histo_stp_sel_ImprimeLibroMayor '" & Format(dFechaDel, gsFormatoFecha) & "','" & Format(dFechaDel, "yyyymmdd") & "','" & Format(dFechaAl, "yyyymmdd") & "','" & IIf(psAge = "", "%%", psAge) & "'," & pnResDet
    sSQL = "exec stp_sel_ImprimeLibroMayor '" & Format(dFechaDel, gsFormatoFecha) & "','" & Format(dFechaDel, "yyyymmdd") & "','" & Format(dFechaAl, "yyyymmdd") & "','" & IIf(psAge = "", "%%", psAge) & "'," & IIf(pnResDet = 2, 1, pnResDet) 'PASI TI-ERS086-2014
End If
'*** FIN PEAC

If oConec.AbreConexion = False Then Exit Function
Set rsCta = oConec.CargaRecordSet(sSQL)
If rsCta.EOF Then
   RaiseEvent BarraClose
   oConec.CierraConexion
   Set oConec = Nothing
   ImprimeLibroMayor = ""
   'Err.Raise 50001, "No se registraron movimientos ...!"
   Exit Function
End If
RaiseEvent BarraClose

RaiseEvent BarraShow(rsCta.RecordCount)
'sTexto = PrnSet("MI", 5)
sImpre = ""
nDebeT = 0: nHaberT = 0
Set oSdo = New NCtasaldo
Do While Not rsCta.EOF
   'Obtenemos el Saldo de la Cuenta al día dFechaDel
   sCtaCod = rsCta!cCtaContCod
   sCtaDes = rsCta!cCtaContDesc
   nSaldo = rsCta!nSaldoIni
   nSaldoIni = nSaldo
   Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage), 0
   nDebe = 0: nHaber = 0
   nDebeD = 0: nHaberD = 0
   If rsCta.EOF Then
      Exit Do
   End If
   sFecha = Mid(rsCta!cMovNro, 1, 8)

   sTipoCta = RTrim(rsCta!cCtaCaracter)

   Do While rsCta!cCtaContCod = sCtaCod
      DoEvents
      RaiseEvent BarraProgress(rsCta.Bookmark, "LIBRO MAYOR", "", "Procesando... ", vbBlue)
      If P Mod 20 = 0 Then
         sImpre = sImpre & sTexto
         sTexto = ""
      End If
      'If Mid(rsCta!cCtaContCod, 1, 6) = "191106" Then
        
        
       ' MsgBox "esto es"
      'End If
      If sFecha <> Mid(rsCta!cMovNro, 1, 8) Then
        'By Capi 071002008
        'ALPA 20090930**********************************************
         'Linea sTexto, CON & Space(12) & "Total Diario  : " & PrnVal(nDebeD, 14, 2, False) & "  " & PrnVal(nHaberD, 14, 2, False) & "  " & PrnVal(nSaldo, 14, 2) & COFF, 2
         
         '*** PEAC 20110606
         'Linea sTexto, Space(12) & "Total Diario  : " & PrnVal(nDebeD, 14, 2, False) & "  " & PrnVal(nHaberD, 14, 2, False) & "  " & PrnVal(nSaldo, 14, 2), 2
         
'         Linea sTexto, CON & Space(12) & "Total Diario  : " & PrnVal(nDebeD, 14, 2, False) & "  " & PrnVal(nHaberD, 14, 2, False) & "  " & PrnVal(nSaldo, 14, 2), 2
         '**********************************************************
         
           '*** PEAC 20110606
'         nLin = nLin + 2
'         nDebeD = 0: nHaberD = 0
'         sFecha = Mid(rsCta!cMovNro, 1, 8)
      End If
      Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage, pnResDet), 0 'pnResDet Agregado 'PASI TI-ERS086-2014
      If sTipoCta = "D" Then
         nSaldo = nSaldo + rsCta!nDebe - rsCta!nHaber
      Else
         nSaldo = nSaldo + rsCta!nHaber - rsCta!nDebe
      End If
      nDebe = nDebe + rsCta!nDebe
      nHaber = nHaber + rsCta!nHaber
      nDebeD = nDebeD + rsCta!nDebe
      nHaberD = nHaberD + rsCta!nHaber
      nDebeT = nDebeT + rsCta!nDebe
      nHaberT = nHaberT + rsCta!nHaber

      'Linea Detalle de Reporte
      'By Capi 07102008
'      Linea sTexto, CON & Mid(rsCta!cMovNro, 7, 2) & "/" & Mid(rsCta!cMovNro, 5, 2) & "/" & Mid(rsCta!cMovNro, 1, 4) & " " & PrnVal(rsCta!nMovCorrela, 5, 0) & "   " & Mid(rsCta!cMovNro, 18, 2) & "-" _
'             & rsCta!cOpeGruCod & Space(4) & PrnVal(rsCta!nDebe, 14, 2, False) & "  " _
'             & PrnVal(rsCta!nHaber, 14, 2, False) & "  " _
'             & PrnVal(nSaldo, 14, 2) & "   " & rsCta!cMovdesc & COFF
'
'ALPA 20090930*****************************
    If pnResDet <> 2 Then 'PASI TI-ERS086-2014
        Linea sTexto, Mid(rsCta!cMovNro, 7, 2) & "/" & Mid(rsCta!cMovNro, 5, 2) & "/" & Mid(rsCta!cMovNro, 1, 4) & " " & PrnVal(rsCta!nMovCorrela, 5, 0) & "   " & Mid(rsCta!cMovNro, 18, 2) & "-" _
                & rsCta!cOpeGruCod & Space(4) & PrnVal(rsCta!nDebe, 14, 2, False) & "  " _
                & PrnVal(rsCta!nHaber, 14, 2, False) & "  " _
                & PrnVal(nSaldo, 14, 2) & "   " & Left(rsCta!cMovdesc, 45)
    End If
' Linea sTexto, CON & Mid(rsCta!cMovNro, 7, 2) & "/" & Mid(rsCta!cMovNro, 5, 2) & "/" & Mid(rsCta!cMovNro, 1, 4) & " " & PrnVal(rsCta!nMovCorrela, 5, 0) & "   " & Mid(rsCta!cMovNro, 18, 2) & "-" _
'             & rsCta!cOpeGruCod & Space(4) & PrnVal(rsCta!nDebe, 14, 2, False) & "  " _
'             & PrnVal(rsCta!nHaber, 14, 2, False) & "  " _
'             & PrnVal(nSaldo, 14, 2) & "   " & Left(rsCta!cMovdesc, 45)
'******************************************
      nLin = nLin + 1
      rsCta.MoveNext
      If rsCta.EOF Then
         Exit Do
      End If
   Loop
   'By Capi 07102008
'   Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage), 0
'   Linea sTexto, CON & Space(12) & "Total Diario  : " & PrnVal(nDebeD, 14, 2, False) & "  " & PrnVal(nHaberD, 14, 2, False) & "  " & PrnVal(nSaldo, 14, 2) & COFF, 2
'   nLin = nLin + 2
'   Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage), 0
'   Linea sTexto, CON & Space(12) & "Total Cuenta  : " & PrnVal(nDebe, 14, 2, False) & "  " & PrnVal(nHaber, 14, 2, False) & "  " & PrnVal(nSaldo, 14, 2) & COFF, 2
'   nLin = nLin + 2
'   Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage) & CON, 0
'   Linea sTexto, Space(12) & "SALDO INICIAL : " & PrnVal(nSaldoIni, 14, 2)
'   Linea sTexto, Space(12) & "SALDO FINAL   : " & PrnVal(nSaldo, 14, 2) & COFF
'   Linea sTexto, "----------------------------------------------------------------------------------------------------------", 2
'ALPA 20090930*****************************
'   Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage), 0
'   Linea sTexto, CON & Space(12) & "Total Diario  : " & PrnVal(nDebeD, 14, 2, False) & "  " & PrnVal(nHaberD, 14, 2, False) & "  " & PrnVal(nSaldo, 14, 2), 2
'   nLin = nLin + 2
'   Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage), 0
'   Linea sTexto, CON & Space(12) & "Total Cuenta  : " & PrnVal(nDebe, 14, 2, False) & "  " & PrnVal(nHaber, 14, 2, False) & "  " & PrnVal(nSaldo, 14, 2), 2
'   nLin = nLin + 2
'   Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage), 0
'   Linea sTexto, Space(12) & "SALDO INICIAL : " & PrnVal(nSaldoIni, 14, 2)
'   Linea sTexto, Space(12) & "SALDO FINAL   : " & PrnVal(nSaldo, 14, 2)
'   Linea sTexto, "----------------------------------------------------------------------------------------------------------", 2

    '*** PEAC 20110606
'   Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage), 0
'   Linea sTexto, Space(12) & "Total Diario  : " & PrnVal(nDebeD, 14, 2, False) & "  " & PrnVal(nHaberD, 14, 2, False) & "  " & PrnVal(nSaldo, 14, 2), 2
'   nLin = nLin + 2
   
   Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage, pnResDet), 0 'pnResDet Agregado 'PASI TI-ERS086-2014
   Linea sTexto, Space(12) & "Total Cuenta  : " & PrnVal(nDebe, 14, 2, False) & "  " & PrnVal(nHaber, 14, 2, False) & "  " & PrnVal(nSaldo, 14, 2), 2
   nLin = nLin + 2
   Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage, pnResDet), 0 'pnResDet Agregado 'PASI TI-ERS086-2014
   Linea sTexto, Space(12) & "SALDO INICIAL : " & PrnVal(nSaldoIni, 14, 2)
   Linea sTexto, Space(12) & "SALDO FINAL   : " & PrnVal(nSaldo, 14, 2)
   Linea sTexto, "----------------------------------------------------------------------------------------------------------", 2
   '******************************************
   'By capi 29112008
   'nLin = nLin + 4
   nLin = nLin + 4

   If Not rsCta.EOF Then
      If nLin > pnLinPage - 8 Then
         sCtaCod = rsCta!cCtaContCod
         sCtaDes = rsCta!cCtaContDesc
         Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage), 0
      Else
         Linea sTexto, "CUENTA: " & Justifica(rsCta!cCtaContCod, 22) & " " & rsCta!cCtaContDesc, 2
         nLin = nLin + 2
      End If
   End If
Loop
Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage, pnResDet), 0 'pnResDet Agregado 'PASI TI-ERS086-2014
'By Capi 07102008
'Linea sTexto, BON & Space(15) & "TOTALES       : " & PrnVal(nDebeT, 16, 2, False) & "  " & PrnVal(nHaberT, 16, 2, False) & BOFF, 2
Linea sTexto, CON & Space(15) & "TOTALES       : " & PrnVal(nDebeT, 16, 2, False) & "  " & PrnVal(nHaberT, 16, 2, False), 2
nLin = nLin + 2
ImprimeLibroMayor = sImpre & sTexto

   RaiseEvent BarraClose
   oConec.CierraConexion
   Set oConec = Nothing
   Set oSdo = Nothing

    Exit Function
ImprimeLibroMayorErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeLibroMayor Method")
End Function

Private Function ImprimeLibroMayorCab(ByRef nLin As Integer, ByRef P As Integer, dFechaDel As Date, dFechaAl As Date, sCtaCod As String, sCtaDes As String, pnLinPage As Integer, Optional pnTpoImpr As Integer = 0) As String 'pnTpoImpr agregado x PASI TI-ERS086-2014
Dim K As Integer, nObjs As Integer
Dim sCabe As String
If nLin > pnLinPage - 6 Then
   'By Capi 07102008
'   Linea sCabe, BON & Cabecera(" L I B R O    M A Y O R ", P, "", 80, , Format(dFechaAl, gsFormatoFechaView), lsEmprLogo), 0
'   Linea sCabe, BON & Centra("Del " & dFechaDel & " al " & dFechaAl, 80) & BOFF, 2
'   Linea sCabe, CON & "FECHA       DOC   ASIENTO            DEBE         HABER        SALDO FINAL   DESCRIPCION" & COFF, 2
'   Linea sCabe, "CUENTA: " & Justifica(sCtaCod, 22) & " " & sCtaDes & BOFF, 2
'ALPA 20090930****************************************
'   Linea sCabe, BON & Cabecera(" L I B R O    M A Y O R ", P, "", 80, , Format(dFechaAl, gsFormatoFechaView), lsEmprLogo), 0
'   Linea sCabe, BON & Centra("Del " & dFechaDel & " al " & dFechaAl, 80), 2
'   Linea sCabe, CON & "FECHA       DOC   ASIENTO            DEBE         HABER        SALDO FINAL   DESCRIPCION", 2
'   Linea sCabe, "CUENTA: " & Justifica(sCtaCod, 22) & " " & sCtaDes, 2
'******************************************************
   Linea sCabe, Cabecera(" L I B R O    M A Y O R ", P, "", 80, , Format(dFechaAl, gsFormatoFechaView), lsEmprLogo), 0
   Linea sCabe, Centra("Del " & dFechaDel & " al " & dFechaAl, 80), 2
   If pnTpoImpr = 2 Then 'PASI TI-ERS086-2014
        Linea sCabe, "                                     DEBE         HABER        SALDO FINAL              ", 2
   Else
        Linea sCabe, "FECHA       DOC   ASIENTO            DEBE         HABER        SALDO FINAL   DESCRIPCION", 2
   End If
   Linea sCabe, "CUENTA: " & Justifica(sCtaCod, 22) & " " & sCtaDes, 2
   '
   nLin = 10
End If
ImprimeLibroMayorCab = sCabe
End Function


Public Function ImprimeMayorCta(psCtaCod As String, psCtaDes As String, pdFechaDel As Date, pdFechaAl As Date, pnImporte As Currency, pnSaldoIniME As Currency, pnSaldoFinalME As Currency, psFiltro As String, pnLinPage As Integer, Optional pbVerPersona As Boolean = False) As String
Dim sTexto As String
Dim nRow As Integer
Dim sMov As String
Dim sDoc As String
Dim N As Integer
Dim nDebe  As Currency, nHaber  As Currency
Dim nDebeD As Currency, nHaberH As Currency
Dim nHaberD As Currency
Dim nSaldo  As Currency, nSaldoIni As Currency
Dim sFecha  As String
Dim rs      As ADODB.Recordset
Dim rsCta   As ADODB.Recordset
Dim nLin    As Integer
Dim P       As Integer
Dim sTipoCta As String

nLin = pnLinPage
P = 0
Dim oCont As New NContAsientos
Dim oSdo  As New NCtasaldo
Set oSdo = Nothing

Set rsCta = oCont.GetMayorCuenta(psCtaCod, Format(pdFechaDel, "yyyymmdd"), Format(pdFechaAl, "yyyymmdd"), pnImporte, psFiltro, , , pbVerPersona)
Set oCont = Nothing
If rsCta.EOF Then
   nSaldo = oSdo.GetCtaSaldo(psCtaCod + "%", Format(pdFechaDel - 1, gsFormatoFecha))
   Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
'   Linea sTexto, BON & "CUENTA CONTABLE : " & Justifica(psCtaCod, 16) & " " & psCtaDes, 2


'   Linea sTexto, CON & Space(87) & "SALDO INICIAL : " & PrnVal(nSaldo, 16, 2)
'   Linea sTexto, Space(87) & "SALDO FINAL   : " & PrnVal(nSaldo, 16, 2) & COFF & BOFF, 1
   
'By Capi 07102008
   Linea sTexto, Space(87) & "SALDO INICIAL : " & PrnVal(nSaldo, 16, 2)
   Linea sTexto, Space(87) & "SALDO FINAL   : " & PrnVal(nSaldo, 16, 2), 1
'
   ImprimeMayorCta = sTexto
   Set oCont = Nothing
   Exit Function
End If
Dim oCta  As New DCtaCont

Set rs = oCta.CargaCtaContClase(psCtaCod)
Set oCta = Nothing
If Not rs.EOF Then
   sTipoCta = Trim(rs!cCtaCaracter)
Else
   sTipoCta = "D"
End If

RaiseEvent BarraShow(rsCta.RecordCount)
sTexto = ""
Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
'Linea sTexto, BON & "CUENTA CONTABLE : " & Justifica(psCtaCod, 16) & " " & psCtaDes & BOFF, 2
nLin = nLin + 2
nDebe = 0: nHaber = 0
N = 1

sFecha = Mid(rsCta!cMovNro, 1, 8)
nSaldo = oSdo.GetCtaSaldo(rsCta!cCtaContCod, Format(pdFechaDel - 1, gsFormatoFecha))
nSaldoIni = nSaldo
Do While Not rsCta.EOF
   DoEvents
   RaiseEvent BarraProgress(rsCta.Bookmark, "Mayor de Cuenta Contable : " & psCtaCod, "", "Generando Reporte... ", vbBlue)
   sMov = rsCta!cMovNro
   If sFecha <> Mid(sMov, 1, 8) Then
      'By Capi 07102008
      Linea sTexto, CON & BON & Space(87) & "TOTAL DIARIO  : " & PrnVal(nDebeD, 16, 2, False) & "  " & PrnVal(nHaberD, 16, 2, False) & "  " & PrnVal(nSaldo, 16, 2) & COFF & BOFF, 2, nLin
      'Linea sTexto, Space(87) & "TOTAL DIARIO  : " & PrnVal(nDebeD, 16, 2, False) & "  " & PrnVal(nHaberD, 16, 2, False) & "  " & PrnVal(nSaldo, 16, 2), 2, nLin
      nDebeD = 0: nHaberD = 0
      sFecha = Mid(rsCta!cMovNro, 1, 8)
   End If
   If sTipoCta = "D" Then
      nSaldo = nSaldo + rsCta!nDebe - rsCta!nHaber
   Else
      nSaldo = nSaldo + rsCta!nHaber - rsCta!nDebe
   End If
   
   Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
   'Linea Detalle de Reporte
   sDoc = Justifica(rsCta!cDocAbrev, 3) & " " & Justifica(rsCta!cDocNro, 20) & " " & Justifica(rsCta!dDocFecha, 10) & " "
   sMov = rsCta!cMovNro
   
   'By Capi 07102008
   Linea sTexto, CON & " " & Mid(sMov, 1, 8) & "-" & Mid(sMov, 9, 6) & " " & Right(RTrim(sMov), 4) & "  " _
          & " " & sDoc & Mid(Replace(Replace(rsCta!cMovdesc, Chr(13), " "), oImpresora.gPrnSaltoLinea, "") & Space(100), 1, 42) & " " _
          & PrnVal(rsCta!nDebe, 16, 2, False) & "  " & PrnVal(rsCta!nHaber, 16, 2, False) _
          & COFF

'    Linea sTexto, " " & Mid(sMov, 1, 8) & "-" & Mid(sMov, 9, 6) & " " & Right(RTrim(sMov), 4) & "  " _
'          & " " & sDoc & Mid(Replace(Replace(rsCta!cMovdesc, Chr(13), " "), oImpresora.gPrnSaltoLinea, "") & Space(100), 1, 42) & " " _
'          & PrnVal(rsCta!nDebe, 16, 2, False) & "  " & PrnVal(rsCta!nHaber, 16, 2, False)
'
   
   nLin = nLin + 1
   nDebeD = nDebeD + rsCta!nDebe
   nDebe = nDebe + rsCta!nDebe
   nHaberD = nHaberD + rsCta!nHaber
   nHaber = nHaber + rsCta!nHaber
   rsCta.MoveNext
   If rsCta.EOF Then
      Exit Do
   End If
   Do While sMov = rsCta!cMovNro
      sDoc = Justifica(rsCta!cDocAbrev, 3) & " " & Justifica(rsCta!cDocNro, 20) & " " & rsCta!dDocFecha & " "
      
      'By Capi 07102008
      Linea sTexto, CON & " " & Space(23) & sDoc & "  " & COFF, , nLin
      'Linea sTexto, " " & Space(23) & sDoc & "  ", , nLin
      
      rsCta.MoveNext
      If rsCta.EOF Then
         Exit Do
      End If
   Loop
   If rsCta.EOF Then
      Exit Do
   End If
Loop

Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
'By Capi 07102008
Linea sTexto, CON & BON & Space(87) & "TOTAL DIARIO  : " & PrnVal(nDebeD, 16, 2, False) & "  " & PrnVal(nHaberD, 16, 2, False) & "  " & PrnVal(nSaldo, 16, 2) & COFF & BOFF, 2
'Linea sTexto, Space(87) & "TOTAL DIARIO  : " & PrnVal(nDebeD, 16, 2, False) & "  " & PrnVal(nHaberD, 16, 2, False) & "  " & PrnVal(nSaldo, 16, 2), 2
nLin = nLin + 2
N = N + 1
Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
'By Capi 07102008
Linea sTexto, CON & " -------------------------" & String(36, "-") & "------------------------------------------------------------------------------------------------"
Linea sTexto, BON & Space(87) & "TOTAL CUENTA  : " & PrnVal(nDebe, 16, 2, False) & "  " & PrnVal(nHaber, 16, 2, False) & BOFF, 2

'Linea sTexto, " -------------------------" & String(36, "-") & "------------------------------------------------------------------------------------------------"
'Linea sTexto, Space(87) & "TOTAL CUENTA  : " & PrnVal(nDebe, 16, 2, False) & "  " & PrnVal(nHaber, 16, 2, False), 2


nLin = nLin + 2

Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
Linea sTexto, Space(87) & "SALDO INICIAL : " & PrnVal(nSaldoIni, 16, 2) & Space(5) & "SALDO INICIAL ME : " & PrnVal(pnSaldoIniME, 16, 2) ' Se Agrego Saldo Inicial ME GITU
Linea sTexto, Space(87) & "SALDO FINAL   : " & PrnVal(nSaldo, 16, 2) & Space(5) & "SALDO FINAL ME   : " & PrnVal(pnSaldoFinalME, 16, 2)  ' Se Agrego Saldo Final ME GITU

'By Capi 07102008
Linea sTexto, COFF & BOFF
nLin = pnLinPage + 1

'By Capi 07102008
Linea sTexto, CON & " =========================" & String(36, "=") & "================================================================================================" & COFF
'Linea sTexto, " =========================" & String(36, "=") & "================================================================================================"
ImprimeMayorCta = sTexto
RaiseEvent BarraClose
rs.Close: Set rs = Nothing
rsCta.Close: Set rsCta = Nothing

End Function

Private Function ImprimeMayorCuentaCab(ByRef nLin As Integer, ByRef P As Integer, dFechaDel As Date, dFechaAl As Date, sCtaCod As String, sCtaDes As String, pnLinPage As Integer) As String
Dim K As Integer, nObjs As Integer
Dim sCabe As String
Dim oConec As New DConecta

If nLin > pnLinPage - 6 Then
   oConec.AbreConexion
   
   'by capi 07102008
   'Linea sCabe, BON & Cabecera(" M A Y O R   D E   C U E N T A S    C O N T A B L E S", P, "", 80, , Format(oConec.GetFechaHoraServer(), gsFormatoFechaView)), 0
   'oConec.CierraConexion
   'Linea sCabe, BON & Centra("Del " & dFechaDel & " al " & dFechaAl, 80), 2
   'Linea sCabe, CON & " ======================" & String(36, "=") & "===================================================================================================="
   'Linea sCabe, "  Número               " & "     DOCUMENTO                       " & "                                                          MOVIMIENTO                         "
   'Linea sCabe, "  de                   " & " ------------------------------------" & " DESCRIPCION                                     --------------------------        S A L D O "
   'Linea sCabe, "  Movimiento           " & " Tipo  Número              Fecha     " & "                                                       DEBE          HABER                   "
   'Linea sCabe, " ----------------------" & String(36, "-") & "----------------------------------------------------------------------------------------------------" & COFF
   'Linea sCabe, BON & "CUENTA CONTABLE : " & Justifica(sCtaCod, 16) & " " & sCtaDes & BOFF, 2
   
   Linea sCabe, Cabecera(" M A Y O R   D E   C U E N T A S    C O N T A B L E S", P, "", 80, , Format(oConec.GetFechaHoraServer(), gsFormatoFechaView)), 0
   oConec.CierraConexion
   Linea sCabe, Centra("Del " & dFechaDel & " al " & dFechaAl, 80), 2
   Linea sCabe, " ======================" & String(36, "=") & "===================================================================================================="
   Linea sCabe, "  Número               " & "     DOCUMENTO                       " & "                                                          MOVIMIENTO                         "
   Linea sCabe, "  de                   " & " ------------------------------------" & " DESCRIPCION                                     --------------------------        S A L D O "
   Linea sCabe, "  Movimiento           " & " Tipo  Número              Fecha     " & "                                                       DEBE          HABER                   "
   Linea sCabe, " ----------------------" & String(36, "-") & "----------------------------------------------------------------------------------------------------"
   Linea sCabe, BON & "CUENTA CONTABLE : " & Justifica(sCtaCod, 16) & " " & sCtaDes, 2
   nLin = 12
End If
ImprimeMayorCuentaCab = sCabe
Set oConec = Nothing
End Function


'************************
'*** NUEVAS FUNCIONES
'************************

Public Function ImprimeRepCtaCol(pnLinPage As Integer) As String
Dim rs As ADODB.Recordset
Dim lsOperacion As String
Dim lnColumna As Integer

Dim sImpre As String
Dim nLin As Integer
Dim nPageNro As Integer
Dim clsRepCtaCol As New DRepCtaColumna
nLin = pnLinPage
sImpre = ""
On Error GoTo ImprimeRepCtaColErr
Set rs = clsRepCtaCol.CargaRepImpresion
If Not rs.EOF Then
   RaiseEvent BarraShow(rs.RecordCount)
   Do While Not rs.EOF
      RaiseEvent BarraProgress(rs.Bookmark, "REPORTES EN BASE A COLUMNAS", "Generando Reporte", "", vbBlue)
      If nLin > pnLinPage - 4 Then
         Linea sImpre, CabeRepo(vsEmpNom, vsAgeNom, "Contabilidad", "", vsFecSis, "Configuración de Reportes de Cuentas Contables por Columnas", _
               String(65, "-"), "", "", nPageNro, 79), 0
         nLin = 6
      End If
      If lsOperacion <> Trim(rs!cOpeCod) Then
          Linea sImpre, ""
          Linea sImpre, BON & Trim(rs!cOpeCod) & " - " & Trim(rs!cOpeDesc) & " " & IIf(Mid(rs!cOpeCod, 3, 1) = "1", "MN", "ME") & " (Anexo : " & Trim(Right(rs!cOpeCod, 1)) & ")" & BOFF
          nLin = nLin + 2
      End If
      If lnColumna <> rs!nNroCol Then
          Linea sImpre, Space(5) & rs!nNroCol & "  " & Trim(rs!cDescCol)
          nLin = nLin + 1
      End If
      Linea sImpre, Space(15) & rs!cCtaContCod & " " & Trim(rs!cCtaContDesc)
      nLin = nLin + 1
      lnColumna = Trim(rs!nNroCol)
      lsOperacion = Trim(rs!cOpeCod)
      rs.MoveNext
    Loop
    RaiseEvent BarraClose
End If
RSClose rs
Set clsRepCtaCol = Nothing
ImprimeRepCtaCol = sImpre
Exit Function
ImprimeRepCtaColErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeRepCtaCol Method")
End Function

Public Function ImprimeCuadroReclasificacion(psTipo As String, pdFecha As Date, pnMoneda As Integer, psClase As String, pnTipCambio As Currency, pnLinPage As Integer, Optional psCtaCont As String, Optional psRep As String = "") As String
'NAGL 202102 Agregó psRep
Dim rs As ADODB.Recordset
Dim rsIntDev As ADODB.Recordset '*** PEAC 20130111
Dim sSQL As String
Dim nLin As Integer
Dim P    As Integer
Dim sTit As String, sTitBarra As String
Dim sTexto As String
Dim nTotDif As Currency
Dim oAju  As New DAjusteCont
Dim nTotE As Currency, nTotC As Currency
Dim lsError As String
'***
Dim lsCuentaCont As String
Dim loAsiento As DCtaCont
On Error GoTo ImprimeCuadroReclasificacionErr

Select Case psTipo
   Case "C"
      sTit = "R E C L A S I F I C A C I O N   D E   C A R T E R A"
      sTitBarra = "RECLASIFICACION DE CARTERA"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteReclasificaCartera(Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio)
   Case "D"
      sTit = "G E N E R A C I O N   D E   I N T E R E S E S   D E V E N G A D O S"
      sTitBarra = "INTERESES DEVENGADOS"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(1, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
   Case "S"
      sTit = "G E N E R A C I O N   D E   I N T E R E S E S   E N   S U S P E N S O"
      sTitBarra = "INTERESES EN SUSPENSO"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(2, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
   Case "G"
      sTit = "R E C L A S I F I C A C I O N   D E   G A R A N T I A S"
      sTitBarra = "RECLASIFICACION DE GARANTIAS"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteReclasificaGarantia(Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio)
   Case "P"
      sTit = "G E N E R A C I O N   D E   P R O V I S I O N E S  D E   C A R T E R A"
      sTitBarra = "PROVISIONES DE CARTERA"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(3, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
   Case "5"
      sTit = "C A P I T A L   D E   C R E D I T O S   C A S T I G A D O S"
      sTitBarra = "CAPITAL DE CREDITOS CASTIGADOS"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(Val(psTipo), Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
   Case "6"
      sTit = "I N T E R E S E S    D E   C R E D I T O S   C A S T I G A D O S"
      sTitBarra = "INTERESES DE CREDITOS CASTIGADOS"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(Val(psTipo), Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
   Case "7"
      sTit = "R E V E R S I O N   D E   I N T E R E S E S   D E V E N G A D O S "
      sTitBarra = "REVERSION DE INTERESES DEVENGADOS"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(Val(psTipo), Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont, True)
   Case "8"
      sTit = "G E N E R A C I O N   D E   P R O V I S I O N E S  D E   C A R T A   F I A N Z A"
      sTitBarra = "PROVISIONES DE CARTA FIANZA"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(Val(psTipo), Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
   Case "9"
      sTit = "G E N E R A C I O N   C A L I F I C A C I O N  D E   C O N T I N G E N T E S"
      sTitBarra = "CALIFICACION DE CREDITOS CONTINGENTES"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(Val(psTipo), Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
   Case "10"
      sTit = "C R E D I T O S   E N    G A R A N T I A    F I N A N C I A M I E N T O"
      sTitBarra = "CREDITOS EN GARANTIA DE FINANCIAMIENTO"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(Val(psTipo), Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
   Case "11"
      sTit = "G E N E R A C I O N   D E   P R O V I S I O N E S  D E   C A R T A   F I A N Z A"
      sTitBarra = "PROVISIONES DE CARTERA"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(11, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    Case "12"
      sTit = "G E N E R A C I O N   D E   A J U S T E  P O R  C A L I F I C A C I O N   D E   C A R T E R A"
      sTitBarra = "AJUSTE DE CALIFICACION DE CARTERA "
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(12, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    Case "13"
      sTit = "G E N E R A C I O N   P O R  R I E S G O  P O N D E R A D O"
      sTitBarra = "AJUSTE POR RIESGO PONDERADO"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(13, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    Case "14"
      sTit = "A J U S T E  D E  C A P T A C I O N E S"
      sTitBarra = "AJUSTE DE CAPTACIONES"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(14, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
      '*** PEAC 20130111
      Set rsIntDev = oAju.AjusteIntDevCaptaciones(14, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    Case "15"
      sTit = "A J U S T E  D E  P R O V I S I O N E S  D E  E J E R C I C I O S  A N T"
      sTitBarra = "AJUSTE DE PROVISIONES DE EJERCICIOS ANTERIORES"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(15, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    Case "16"
      sTit = "C O M I S I O N E S  D E  C R E D I T O S  C A S T I G A D O S"
      sTitBarra = "COMISIONES DE CREDITOS CASTIGADOS"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(16, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    'ALPA 20090529*******************************
    Case "17"
      sTit = "REPORTE DE CIERRE DE SALDO DE COFIDE Y AGROBANCO"
      sTitBarra = "CIERRE DE SALDO DE COFIDE Y AGROBANCO"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(17, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    '********************************************
    'ALPA 20090529*******************************
    Case "18"
      sTit = "A J U S T E   D E   I N T E R E S E S   D I F E R I D O S   P A G A D O S"
      sTitBarra = "AJUSTE DE INTERESES DIFERIDOS"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(18, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    '********************************************
    Case "19"
      sTit = "A J U S T E   D E    G A R A N T I A S   DE  C A R T A  F I A N Z A S"
      sTitBarra = "AJUSTE DE GARANTIAS DE CARTAS FIANZAS"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(19, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    '********************************************

 'MADM 20110805
    Case "20"
      'sTit = "A J U S T E   D E   S A L D O S  D E  G A R A N T I A S   DE  C A R T A  F I A N Z A S"
      sTit = "A J U S T E   D E   S A L D O S  D E  C A R T A  F I A N Z A S" 'EJVG20130322
      'sTitBarra = "AJUSTE DE GARANTIAS DE CARTAS FIANZAS"
      sTitBarra = "AJUSTE DE SALDOS DE CARTAS FIANZAS"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(20, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    '********************************************
    'ALPA 20111005
    Case "21"
    sTit = "A J U S T E   D E   I N T E R E S E S   D I F E R I D O S   T R A N S F E R I D O S"
      sTitBarra = "AJUSTE DE INTERESES DIFERIDOS"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(21, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    '********************************************
'END MADM
    'JUEZ 20130116 *******************************************************************
    Case "22"
      sTit = "G A S T O S    D E   C R E D I T O S   C A S T I G A D O S"
      sTitBarra = "GASTOS DE CREDITOS CASTIGADOS"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(Val(psTipo), Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    'END JUEZ ************************************************************************
    'EJVG20130307 ***
    Case "23"
      sTit = "A S I E N T O   D E   C O M I S I O N   C A R T A   F I A N Z A"
      sTitBarra = "COMISION CARTA FIANZA"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(Val(psTipo), Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    'END EJVG ******
     'ALPA20131129 ***
    Case "24"
      sTit = "A S I E N T O   D E   I N T E R E S E S   D I F E R I D O S   -   S A L D O S     F I N A L"
      sTitBarra = "INTERES DIFERIDOS"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(Val(psTipo), Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    'END ************
    'ALPA20140925 ***
    Case "25"
      sTit = "A S I E N T O   D E   C O M P R A   DE   C A R T E R A"
      sTitBarra = "COMPRA DE CARTERA"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(Val(psTipo), Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    'END ************
    Case "26"
      sTit = "A S I E N T O   D E   R E C L A S I F I C A C I O N  D E  P R O V I S I O N E S  P R O C I C L I C A S "
      sTitBarra = "RECLASIFICACION DE PROVISIONES PROCICLICAS"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(Val(psTipo), Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    'END ************
    'ALPA 20160222***********************
    Case "27"
      sTit = "A S I E N T O   D E   D I F E R I D O S    D E    R E C L A S I F I C A C I O N    D E    C R E D I T O S    R E F I N A N C I A D O S "
      sTitBarra = "DIFERIDOS DE CREDITOS AMPLIADOS"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(Val(psTipo), Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    'END ********************************
    'PASI 20170412 **********************
    Case "28"
      sTit = "R E V E R S I O N   D E   I N T E R E S E S   D E V E N G A D O S   X   Z O N A   I N U N D A D A"
      sTitBarra = "REVERSION DE INTERESES DEVENGADOS X ZONA INUNDADA"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(Val(psTipo), Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont, True)
    'PASI END****************************
    'NAGL 202007 Según ACTA N°049-2020************
    Case "29"
      sTit = "REPORTE CIERRE DE CARTERA FAE-MYPE"
      sTitBarra = "CIERRE DE SALDO DE FAE-MYPE"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(29, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    '********************************************
    'NAGL 202007 Según ACTA N°049-2020************
    Case "30"
      sTit = "REPORTE CIERRE DE CARTERA REACTIVA"
      sTitBarra = "CIERRE DE SALDO DE REACTIVA"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(30, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    '********************************************
    'NAGL 202007 Según ACTA N°049-2020************
    Case "31"
      sTit = "REPORTE CIERRE DE CARTERA REPROGRAMADA"
      sTitBarra = "CIERRE DE SALDO DE CARTERA REPROGRAMADA"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(31, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont)
    '********************************************
    'NAGL 202102 Según ACTA N°017-2021************
    Case "34"
      sTit = "REPORTE CIERRE DE PROVISIONES DE CARTERA REPROGRAMADA COVID"
      sTitBarra = "CIERRE DE SALDO PROVISIONES COVID CARTERA REPROGRAMADA"
      RaiseEvent BarraShow(1)
      RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
      Set rs = oAju.AjusteInteresesColocaciones(34, Format(pdFecha, gsFormatoFecha), pnMoneda, psClase, pnTipCambio, psCtaCont, , , , psRep)
    '********************************************
End Select
RaiseEvent BarraClose
If rs.EOF Then
    'ALPA20140129
    MsgBox "No existen diferencias entre Estadísticas y Saldos Contables", vbInformation, "Aviso"
    ImprimeCuadroReclasificacion = ""
    Exit Function
   'Err.Raise "50001", "NContImprimir: ImprimeCuadroReclasificacion", "No existen diferencias entre Estadísticas y Saldos Contables "
Else

   Set loAsiento = New DCtaCont
   nLin = pnLinPage
   P = 0
   If psTipo = "29" Or psTipo = "30" Or psTipo = "31" Or psTipo = "34" Then 'Fae_Mype Reactiva y Cart. Reprogramada
   'NAGL 202102 Agregó psTipo = "34"
         sTexto = ImprimeCuadroOtrosProcesos(psTipo, rs, sTitBarra, sTit, pnLinPage, loAsiento) 'NAGL 20200804 Según ACTA N°049-2020
         'NAGL 202102 Agregó psTipo
   Else 'Condicional agregado by NAGL 20200804
    RaiseEvent BarraShow(rs.RecordCount)
    Do While Not rs.EOF
       RaiseEvent BarraProgress(rs.Bookmark, sTitBarra, "", "Generando Cuadro...", vbBlue)
       If nLin > pnLinPage - 6 Then
          Linea sTexto, Cabecera(sTit, P, "", 120, , , gsNomCmac), 0
          Linea sTexto, "Cuenta Contable                 S.Estadistico          S.Contable          Diferencia"
          Linea sTexto, "-------------------------------------------------------------------------------------"
          nLin = 5
       End If
       gnImporte = nVal(Format(rs!nSaldo, "#.00"))
       lsCuentaCont = IIf(IsNull(rs!Cta1), rs!cta2, rs!Cta1)
       On Error Resume Next
       lsError = loAsiento.ExisteCuentaCad(Trim(lsCuentaCont), True)
       
       If lsError <> "" Then
         Linea sTexto, loAsiento.ExisteCuentaCad(Trim(lsCuentaCont), True) & lsCuentaCont
       End If
       On Error GoTo ImprimeCuadroReclasificacionErr
       If (psTipo = "12" Or psTipo = "13") Then
         Linea sTexto, Justifica(IIf(IsNull(rs!Cta1), rs!cta2, rs!Cta1), 22) & "     " & PrnVal(gnImporte, 18, 2) & "  " & PrnVal(rs!nCtaSaldoImporte, 18, 2) & "  " & PrnVal(rs!nSaldo - rs!nCtaSaldoImporte, 18, 2)
         nTotDif = nTotDif + (gnImporte - rs!nCtaSaldoImporte)
       Else
         Linea sTexto, Justifica(IIf(IsNull(rs!Cta1), rs!cta2, rs!Cta1), 22) & "     " & PrnVal(gnImporte, 18, 2) & "  " & PrnVal(rs!nCtaSaldoImporte, 18, 2) & "  " & PrnVal(rs!nCtaSaldoImporte - rs!nSaldo, 18, 2)
         nTotDif = nTotDif + (rs!nCtaSaldoImporte - gnImporte)
       End If
       nTotE = nTotE + gnImporte
       nTotC = nTotC + rs!nCtaSaldoImporte
       nLin = nLin + 1
       rs.MoveNext
    Loop
    Linea sTexto, "-------------------------------------------------------------------------------------"
    Linea sTexto, "      TOTAL DIFERENCIA     " & PrnVal(nTotE, 18, 2) & "  " & PrnVal(nTotC, 18, 2) & "  " & PrnVal(nTotDif, 18, 2)
   End If
   
   '*** PEAC 20130111 - MUESTRA LOS AJUSTES DE CAPTACIONES INTERESES DEVENGADOS
    If psTipo = "14" Then
        If Not rsIntDev.EOF Then
            nTotE = 0: nTotC = 0: nTotDif = 0
            nLin = pnLinPage
            
            Do While Not rsIntDev.EOF
                RaiseEvent BarraProgress(rsIntDev.Bookmark, "Ajuste de Captaciones Intereses Devengados", "", "Generando Cuadro...", vbBlue)
                If nLin > pnLinPage - 6 Then
                    Linea sTexto, Cabecera(sTit, P, "", 120, , , gsNomCmac), 0
                    Linea sTexto, Space(40) + "I N T E R E S E S   D E V E N G A D O S"
                    Linea sTexto, "Cuenta Contable                 S.Estadistico          S.Contable          Diferencia"
                    Linea sTexto, "-------------------------------------------------------------------------------------"
                    nLin = 5
                End If
                gnImporte = nVal(Format(rsIntDev!nSaldo, "#.00"))
                lsCuentaCont = IIf(IsNull(rsIntDev!Cta1), rsIntDev!cta2, rsIntDev!Cta1)
                On Error Resume Next
                lsError = loAsiento.ExisteCuentaCad(Trim(lsCuentaCont), True)
                
                If lsError <> "" Then
                    Linea sTexto, loAsiento.ExisteCuentaCad(Trim(lsCuentaCont), True) & lsCuentaCont
                End If
                On Error GoTo ImprimeCuadroReclasificacionErr
                Linea sTexto, Justifica(IIf(IsNull(rsIntDev!Cta1), rsIntDev!cta2, rsIntDev!Cta1), 22) & "     " & PrnVal(gnImporte, 18, 2) & "  " & PrnVal(rsIntDev!nCtaSaldoImporte, 18, 2) & "  " & PrnVal(rsIntDev!nCtaSaldoImporte - rsIntDev!nSaldo, 18, 2)
                nTotDif = nTotDif + (rsIntDev!nCtaSaldoImporte - gnImporte)
                
                nTotE = nTotE + gnImporte
                nTotC = nTotC + rsIntDev!nCtaSaldoImporte
                nLin = nLin + 1
                rsIntDev.MoveNext
            Loop
            Linea sTexto, "-------------------------------------------------------------------------------------"
            Linea sTexto, "      TOTAL DIFERENCIA     " & PrnVal(nTotE, 18, 2) & "  " & PrnVal(nTotC, 18, 2) & "  " & PrnVal(nTotDif, 18, 2)
        End If
    End If
   '*** FIN PEAC
   
   ImprimeCuadroReclasificacion = sTexto
   RaiseEvent BarraClose
   Set loAsiento = Nothing
End If
Exit Function
ImprimeCuadroReclasificacionErr:
    RaiseEvent BarraClose
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeCuadroReclasificacion Method")
End Function


Public Function ImprimeFlujoCtasBancos(ByVal pnColPage As Integer, ByVal pnLinPage As Integer, ByVal psOpecod As String, _
                            ByVal psCtaIfCodDesde As String, ByVal psCtaIfCodHasta As String, _
                            ByVal pdDesde As Date, ByVal pdHasta As Date) As String

Dim N As Integer
Dim P As Integer
Dim nLin As Integer
Dim lnTotD As Currency, lnTotH As Currency, lnSaldo As Currency
Dim lbOk As Boolean
Dim lsCond As String
Dim lsMovNro As String
Dim lsDocNro As String
Dim lsCtaIFCod As String
Dim lnImporte As Currency
Dim lsTexto As String
Dim lsGlosa As String
Dim rs As ADODB.Recordset
Dim lsPersCod As String
Dim lsSimbolo As String

Dim oCont As NCajaCtaIF
Set oCont = New NCajaCtaIF
Set rs = New ADODB.Recordset
On Error GoTo ErrRep
nLin = pnLinPage
lsTexto = ""
RaiseEvent BarraShow(1)
RaiseEvent BarraProgress(0, "Generando Flujo de Bancos", "Cargando Datos...", "", vbBlue)

'''lsSimbolo = IIf(Mid(psOpecod, 3, 1) = gMonedaNacional, "S/.", "US$") 'MARG ERS044-2016
lsSimbolo = IIf(Mid(psOpecod, 3, 1) = gMonedaNacional, gcPEN_SIMBOLO, "US$")
Set rs = oCont.GetFlujosCtaIF(psOpecod, psCtaIfCodDesde, psCtaIfCodHasta, pdDesde, pdHasta)
RaiseEvent BarraProgress(1, "Generando Flujo de Bancos", "Cargando Datos...", "", vbBlue)
RaiseEvent BarraClose

If rs.EOF And rs.BOF Then
    rs.Close: Set rs = Nothing
    ImprimeFlujoCtasBancos = ""
    Exit Function
End If
RaiseEvent BarraShow(rs.RecordCount)
lsCtaIFCod = ""
lnImporte = 0
lnTotD = 0
lnTotH = 0
Do While Not rs.EOF
    If nLin >= 56 Then
        nLin = pnLinPage
    End If
    CabeceraFlujo lsTexto, nLin, P, pnLinPage, pnColPage, lsSimbolo, pdDesde, pdHasta
    If rs!cMovNro = lsMovNro And IIf(Mid(psOpecod, 3, 1) = gMonedaExtranjera, rs!nMovMEImporte, rs!nMovImporte) = lnImporte Then
       Linea lsTexto, CON & Space(64) & ImpreFormat(rs!Documento, 20) & " " & COFF
    Else
       If rs!cPerscod + rs!cCtaIFCod <> lsPersCod + lsCtaIFCod Then
           If lnTotD > 0 Or lnTotH > 0 Then
               Linea lsTexto, CON & " ------------------------------------------------------------------------------------------------------------------------------------"
               nLin = nLin + 1
               Linea lsTexto, BON & Space(75) & "TOTALES  " & ImpreFormat(lnTotD, 14, 2, True) & ImpreFormat(lnTotH, 14, 2, True) & BOFF & COFF, 2
               nLin = nLin + 2
           End If
           lnTotD = 0
           lnTotH = 0
           Linea lsTexto, BON & CON & " CUENTA : [" & rs!cCtaIFCod & "] " & rs!cNomCtaIF & COFF & BOFF
           nLin = nLin + 1
           lnTotD = 0: lnTotH = 0: lnSaldo = 0
           Linea lsTexto, CON & " SALDO ANTERIOR..." & Space(100) & " " & Right(Space(14) & Format(IIf(Mid(psOpecod, 3, 1) = gMonedaExtranjera, rs!nSaldoIniME, rs!nSaldoIni), "#,#0.00"), 13) & COFF
           If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
                lnSaldo = rs!nSaldoIniME
           Else
                lnSaldo = rs!nSaldoIni
           End If
           nLin = nLin + 1
       End If
       lsMovNro = rs!cMovNro
       If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
           lnImporte = rs!nMovMEImporte
           lnTotD = lnTotD + rs!DebeME
           lnTotH = lnTotH + rs!HaberME
       Else
           lnImporte = rs!nMovImporte
           lnTotD = lnTotD + rs!Debe
           lnTotH = lnTotH + rs!Haber
       End If
       lnSaldo = lnSaldo + lnImporte
       lsGlosa = rs!cMovdesc
       
       Linea lsTexto, CON & " " & Mid(lsMovNro, 1, 8) & "-" & Mid(lsMovNro, 9, 6) & "-" & Right(lsMovNro, 4) & ImpreFormat(Replace(lsGlosa, Chr(13), ""), 40) & " " _
               & ImpreFormat(rs!Documento, 17) & " " _
               & ImpreFormat(IIf(Mid(psOpecod, 3, 1) = gMonedaExtranjera, rs!DebeME, rs!Debe), 14, 2, True) & ImpreFormat(IIf(Mid(psOpecod, 3, 1) = gMonedaExtranjera, rs!HaberME, rs!Haber), 14, 2, True) & " " _
               & ImpreFormat(lnSaldo, 10, 2, True) & COFF
   
       nLin = nLin + 1
    End If
    RaiseEvent BarraProgress(rs.Bookmark, "Generando Flujo de Bancos", Trim(rs!cpersNombre) & " [" & rs!cCtaIFCod + "]", "", vbBlue)
    lsCtaIFCod = rs!cCtaIFCod
    lsPersCod = rs!cPerscod
    rs.MoveNext
Loop
If lnTotD > 0 Or lnTotH > 0 Then
    Linea lsTexto, CON & " ------------------------------------------------------------------------------------------------------------------------------------"
    nLin = nLin + 1
    Linea lsTexto, BON & Space(75) & "TOTALES  " & ImpreFormat(lnTotD, 14, 2, True) & ImpreFormat(lnTotH, 14, 2, True) & BOFF & COFF, 2
    nLin = nLin + 2
End If
Linea lsTexto, CON & " ====================================================================================================================================" & COFF, 2
ImprimeFlujoCtasBancos = lsTexto
rs.Close
RaiseEvent BarraClose
Set rs = Nothing
Exit Function
ErrRep:
    ImprimeFlujoCtasBancos = ""
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
Private Function CabeceraFlujo(ByRef lsTexto As String, ByRef nLin As Integer, ByRef P As Integer, _
            ByVal pnLinPage As Integer, pnColPage As Integer, ByVal psSimbolo As String, ByVal pdDesde As Date, ByVal pdHasta As Date) As String
If nLin >= pnLinPage - 7 Then
   Linea lsTexto, Cabecera(" FLUJO  DE  ENTIDADES  FINANCIERAS ", P, psSimbolo, pnColPage, , , gsNomCmac)
   Linea lsTexto, Centra("( DEL " & pdDesde & " AL " & pdHasta & ")", pnColPage) & BOFF
   Linea lsTexto, PrnSet("C+")
   nLin = 11
   Linea lsTexto, " ===================================================================================================================================="
   Linea lsTexto, " Número de             CONCEPTO                                       DOCUMENTO                        MOVIMIENTO             SALDOS "
   Linea lsTexto, " Movimiento                                                          Tipo  Número                 DEBE           HABER               "
   Linea lsTexto, " ------------------------------------------------------------------------------------------------------------------------------------" & PrnSet("B-")
End If
End Function
Public Function ImprimeSaldosCtaIf(ByVal pnColPage As Integer, ByVal pnLinPage As Integer, ByVal psOpecod As String, _
                                ByVal pdFecha As Date) As String
Dim rs As ADODB.Recordset
Dim nLin  As Integer
Dim oCtaIf As NCajaCtaIF
Dim lsTexto As String
Dim lsSimbolo As String
Dim P As Integer
Set oCtaIf = New NCajaCtaIF
Set rs = New ADODB.Recordset
lsTexto = ""
lsSimbolo = IIf(Mid(psOpecod, 3, 1) = gMonedaNacional, gcPEN_SIMBOLO, "$.")
Set rs = oCtaIf.GetRepSaldosCtaIf(psOpecod, pdFecha, , gEstadoCtaIFActiva, Mid(psOpecod, 3, 1))
nLin = pnLinPage
If Not rs.EOF And Not rs.BOF Then
    RaiseEvent BarraShow(rs.RecordCount)
    Do While Not rs.EOF
       If Len(rs!cCtaIFCod) = 15 Then
          Linea lsTexto, ""
          nLin = nLin + 1
       End If
       RaiseEvent BarraProgress(rs.Bookmark, "Generando Saldos de Cuentas", "", "", vbBlue)
       CabeceraRepoSaldo lsTexto, nLin, P, pnLinPage, pnColPage, lsSimbolo, pdFecha, psOpecod
       Linea lsTexto, CON & " " & Mid(rs!cCtaIFCod & Space(20), 1, 25) & "   " & Mid(Trim(rs!cBancoDesc) & " " & rs!cCtaIFDesc & Space(70), 1, 55) & " " _
              & Right(Space(18) & Format(rs!Debe, "#,#0.00"), 18) & " " & Right(Space(10) & Format(rs!Haber, "#,#0.00"), 18) & " " & COFF
       nLin = nLin + 1
       rs.MoveNext
    Loop
    Linea lsTexto, CON & " ====================================================================================================================================" & COFF, 2
    RaiseEvent BarraClose
End If
rs.Close: Set rs = Nothing
ImprimeSaldosCtaIf = lsTexto
End Function
Private Sub CabeceraRepoSaldo(ByRef lsTexto As String, ByRef nLin As Integer, ByRef P As Integer, _
                        ByVal pnLinPage As Integer, pnColPage As Integer, ByVal psSimbolo As String, ByVal pdHasta As Date, ByVal psOpecod As String)
If nLin > pnLinPage - 6 Then
   Linea lsTexto, Cabecera(" SALDOS  DE  ENTIDADES  FINANCIERAS ", P, psSimbolo, pnColPage, , , gsNomCmac)
   Linea lsTexto, Centra("( AL " & pdHasta & " )", pnColPage) & BOFF
   Linea lsTexto, CON
   nLin = 11
   Linea lsTexto, " ============================================================================================================================"
   Linea lsTexto, "   ENTIDAD FINANCIERA                                                                                 S A L D O S   M. " & IIf(Mid(psOpecod, 3, 1) = gMonedaNacional, "N.", "E.")
   Linea lsTexto, "                                                                                                    DEBE              HABER  "
   Linea lsTexto, " ----------------------------------------------------------------------------------------------------------------------------"
End If
End Sub

Public Function ImprimeComprobanteRetencion(Optional pnMovNro As Long = 0, Optional psDocNro As String = "") As String
    Dim sqlP As String
    Dim sqlD As String
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    Dim rsD As ADODB.Recordset
    Set rsD = New ADODB.Recordset
    Dim oConst As NConstSistemas
    Set oConst = New NConstSistemas
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim lsCtaReten As String
    Dim lsCtaIGV As String
    
    Dim lsCadena As String
    Dim lsMargen As String
    
    Dim lsTpoDoc As String * 10
    Dim lsDocSerie As String * 10
    Dim lsDocNro As String * 12
    Dim lsDocFecEmision As String * 10
    Dim lsDocMonto As String * 15
    Dim lsRetenMonto As String * 10
    Dim lnDocMontoTotal As Currency
    
    Dim lnCont As Long
    
    lsCtaReten = oConst.LeeConstSistema(gConstSistCtaRetencion6Porcent)
    lsCtaIGV = oConst.LeeConstSistema(gConstCtaIGV)
    
    lsCadena = ""
    
    lsMargen = Space(0)
    
    oCon.AbreConexion
    
    If pnMovNro <> 0 Then
        sqlP = " Select cPersNombre, M.cMovNro, IsNull((Select cPersIDnro From PersID PDI Where cPersIDTpo = " & PersIdTipo.gPersIdRUC & " And cPersCod = PE.cPersCod),'') ID, Abs(MC.nMovImporte) RetenMonto From Mov M" _
             & " Inner Join MovGasto MG On M.nMovNro = MG.nMovNro" _
             & " Inner Join Persona PE On MG.cPersCod = PE.cPersCod" _
             & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro And MD.nDocTpo = " & TpoDoc.TpoDocAlmacenComprdeRetencion6 & "" _
             & " Inner Join MovCta MC On M.nMovNro = MC.nMovNro" _
             & " Where M.nMovNro = " & pnMovNro & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & " And M.nMovEstado = " & MovEstado.gMovEstContabMovContable & " And MC.cCtaContCod = '" & lsCtaReten & "'"
        Set rsP = oCon.CargaRecordSet(sqlP)

        sqlD = " Select D.cDocDesc, RMD.dDocFecha, Case When PATINDEX('%-%',cDocNro) = 0 Then '' Else Left(cDocNro,PATINDEX('%-%',cDocNro)-1) End Serie," _
             & " Case When PATINDEX('%-%',cDocNro) = 0 Then cDocNro Else Substring(cDocNro,PATINDEX('%-%',cDocNro) + 1,50)  End Nro," _
             & " ABs(IsNull(RME.nMovMEImporte,RMC.nMovImporte)) nMovImporte, Case When RME.nMovMEImporte Is Null Then " & Moneda.gMonedaNacional & " Else " & MovEstado.gMovEstContabMovContable & " End Moneda" _
             & " From MovRef MR" _
             & " Inner Join Mov RM On MR.nMovNroRef = RM.nMovNro And RM.nMovFlag = " & MovFlag.gMovFlagVigente & " And RM.nMovEstado = " & MovEstado.gMovEstContabMovContable & "" _
             & " Inner Join MovDoc RMD On RM.nMovNro = RMD.nMovNro" _
             & " Inner Join MovCta RMC On RM.nMovNro = RMC.nMovNro" _
             & " Left  Join MovME RME On RMC.nMovNro = RME.nMovNro And RMC.nMovItem = RME.nMovItem" _
             & " Inner Join Documento D On D.nDocTpo = RMD.nDocTpo" _
             & " Where MR.nMovNro = " & pnMovNro & " And RMD.nDocTpo In (Select nDocTpo From DocImpuesto where cCtaContCod = '" & lsCtaIGV & "')" _
             & " And RMC.nMovImporte < 0"
        Set rsD = oCon.CargaRecordSet(sqlD)
        
        
        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & lsMargen & Space(15) & PstaNombre(rsP!cpersNombre, True) & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & lsMargen & Space(15) & rsP!ID & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & lsMargen & Space(15) & GetFechaMov(rsP!cMovNro, True) & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        
        lnCont = 0
        lnDocMontoTotal = 0
        
        While Not rsD.EOF
            lnCont = lnCont + 1
            lsTpoDoc = rsD!cDocDesc
            lsDocSerie = rsD!Serie
            lsDocNro = rsD!Nro
            lsDocFecEmision = Format(rsD!dDocFecha, gsFormatoFechaView)
            If rsD!Moneda = Moneda.gMonedaNacional Then
                '''RSet lsDocMonto = "S/. " & Format(rsD!nMovImporte, "#,##0.00") 'MARG ERS044-2016
                RSet lsDocMonto = gcPEN_SIMBOLO & " " & Format(rsD!nMovImporte, "#,##0.00")
            Else
                RSet lsDocMonto = "U$. " & Format(rsD!nMovImporte, "#,##0.00")
            End If
            lnDocMontoTotal = lnDocMontoTotal + rsD!nMovImporte
            
            If lnCont = 1 Then
                '''RSet lsRetenMonto = "S/. " & Format(rsP!RetenMonto, "#,##0.00") 'MARG ERS044-2016
                RSet lsRetenMonto = gcPEN_SIMBOLO & " " & Format(rsP!RetenMonto, "#,##0.00") 'MARG ERS044-2016
            Else
                lsRetenMonto = ""
            End If
            
            lsCadena = lsCadena & lsMargen & lsTpoDoc & "  " & lsDocSerie & "  " & lsDocNro & "       " & lsDocFecEmision & "       " & lsDocMonto & "        " & lsRetenMonto & oImpresora.gPrnSaltoLinea
            
            rsD.MoveNext
        Wend
        
        While lnCont < 25
            lnCont = lnCont + 1
            lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        Wend
            
        lsTpoDoc = ""
        lsDocSerie = ""
        lsDocNro = ""
        lsDocFecEmision = ""
        rsD.MovePrevious
        If rsD!Moneda = Moneda.gMonedaNacional Then
            '''RSet lsDocMonto = gcMN & " " & Format(lnDocMontoTotal, "#,##0.00") 'MARG ERS044-2016
            RSet lsDocMonto = gcPEN_SIMBOLO & " " & Format(lnDocMontoTotal, "#,##0.00") 'MARG ERS044-2016
        Else
            RSet lsDocMonto = gcME & " " & Format(lnDocMontoTotal, "#,##0.00")
        End If
        '''RSet lsRetenMonto = "S/. " & Format(rsP!RetenMonto, "#,##0.00") 'MARG ERS044-2016
        RSet lsRetenMonto = gcPEN_SIMBOLO & " " & Format(rsP!RetenMonto, "#,##0.00") 'MARG ERS044-2016
        
        lsCadena = lsCadena & lsMargen & lsTpoDoc & "  " & lsDocSerie & "  " & lsDocNro & "       " & lsDocFecEmision & "     " & lsDocMonto & "        " & lsRetenMonto & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & lsMargen & "  " & ConvNumLet(rsP!RetenMonto, True, False, gMonedaNacional) & oImpresora.gPrnSaltoLinea
    
    ElseIf psDocNro <> "" Then
    
    
    End If
    
    ImprimeComprobanteRetencion = lsCadena
    
End Function


Public Function ImprimeSaldoProvision(psSaldoProv As String, pssaldoCanc As String, pssaldo As String, lsfechadel As String, lsfechaal As String) As String
Dim sTexto As String
Dim nRow As Integer
Dim sMov As String
Dim sDoc As String
Dim N As Integer
Dim nDebe  As Currency, nHaber  As Currency
Dim nDebeD As Currency, nHaberH As Currency
Dim nHaberD As Currency
Dim nSaldo  As Currency, nSaldoIni As Currency
Dim sFecha  As String
Dim rs      As ADODB.Recordset
Dim rsCta   As ADODB.Recordset
Dim nLin    As Integer
Dim P       As Integer
Dim sTipoCta As String

nLin = 1
P = 0
Dim oCont As New NContAsientos
Dim oSdo  As New NCtasaldo
Set oSdo = Nothing

sTexto = ""

Linea sTexto, CON & BON & Space(20) & "  SALDO DE PROVISION DE PAGO A PROVEEDORES DEL DIA : " & lsfechadel & "  AL  :   " & lsfechaal & COFF & BOFF, 2, nLin
Linea sTexto, CON & BON & Space(10) & "----------------------------------------------------------------------------" & COFF & BOFF, 2, nLin
Linea sTexto, CON & BON & Space(20) & "Saldo Provisionado         Saldo Cancelado         Saldo  " & COFF & BOFF, 2, nLin
Linea sTexto, CON & BON & Space(10) & "----------------------------------------------------------------------------" & COFF & BOFF, 2, nLin
Linea sTexto, CON & BON & Space(20) & psSaldoProv & "                    " & pssaldoCanc & "                " & pssaldo & "          " & COFF & BOFF, 2, nLin
ImprimeSaldoProvision = sTexto

End Function


Public Function ImprimeAsientoContableNew(ByVal psMovNro As String, ByVal pnLinPage As Integer, ByVal pnColPage As Integer, Optional ByVal psTit1 As String, Optional ByVal psCabe1 As String, Optional ByVal psPiePag As String = "9", Optional ByVal psOtraFirma As String = "", Optional psNomCmact As String = "", Optional pnMonto As Currency, Optional psMoneda As Integer) As String
    On Error GoTo ImprimeAsientoContableErr
    Dim sSQL As String
    Dim rs       As New ADODB.Recordset
    Dim rsDoc       As New ADODB.Recordset
    Dim nTot     As Currency
    Dim nTotH    As Currency
    Dim sDoc     As String
    Dim sTexto   As String
    Dim sAsiento As String
    Dim nLi      As Integer
    Dim sCabecera As String
    Dim lnCambioFijo As Currency
    Dim lnCambioVenta As Currency
    Dim lnCambioCompra As Currency
    Dim lnCambioVentaSBS As Currency
    Dim lnCambioCompraSBS As Currency
    Dim CON As String
    Dim COFF As String
    Dim lsMovBilletaje As String
    
    Dim lsDesObjeto As String
    Dim lsFecha    As String
    Dim oConect As DConecta
    Dim lsGlosa As String
    Dim lsMovNro As String
    Dim lsOpeCod As String
    Dim sDestino As String
    Dim sOrigen As String
    Dim oDGeneral As DGeneral
    Dim BON As String
    Dim BOFF As String
    Dim oTipoCamb As nTipoCambio
    
    Set oDGeneral = New DGeneral
    Set oTipoCamb = New nTipoCambio
    
    Set oConect = New DConecta
    
    If oConect.AbreConexion = False Then Exit Function
    
    sCabecera = ""
    sAsiento = ""
    nLi = 0
    If psMovNro = "" Then
       Exit Function
    End If
       lsFecha = GetFechaMov(psMovNro, True)
       sSQL = "SELECT * FROM Mov WHERE cMovNro = '" & psMovNro & "'"
       Set rs = oConect.CargaRecordSet(sSQL)
       If rs.EOF Then
          Exit Function
       End If
       lsGlosa = Replace(rs!cMovdesc, Chr(13) & oImpresora.gPrnSaltoLinea, " ")
       lsMovNro = psMovNro
       lsOpeCod = rs!cOpeCod
       If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Or Left(lsOpeCod, 4) = Left(gOpeME, 4) Or rs!cOpeCod = gContAjusteTipoCambio Then
         lnCambioFijo = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCFijoMes)
         lnCambioVenta = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCVenta)
         lnCambioCompra = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCCompra)
         lnCambioVentaSBS = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCPondVenta)
         lnCambioCompraSBS = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCPonderado)
       End If
       
       If rs!nMovFlag = gMovFlagDeExtorno Then
            psTit1 = " A S I E N T O  C O N T A B L E   D E   E X T O R N O"
       ElseIf rs!nMovFlag = gMovFlagEliminado Then
            psTit1 = " **** ASIENTO ELIMINADO NO VALIDO PARA CONTABILIDAD **** "
       End If
       'sAsiento = sAsiento & ImpreCabAsiento(pnColPage, CDate(lsFecha), gsNomCmac, lsOpeCod, lsMovNro, psTit1, True, , , nLi) & oImpresora.gPrnSaltoLinea
       'sAsiento = sAsiento & ImpreCabAsientoNew(pnColPage, CDate(lsFecha), gsNomCmac, lsOpeCod, lsMovNro, psTit1, True, , , nLi) & oImpresora.gPrnSaltoLinea
       If psCabe1 <> "" Then
          sAsiento = sAsiento & Space(10) & lsGlosa & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sAsiento = sAsiento & psCabe1 & oImpresora.gPrnSaltoLinea
          nLi = nLi + 4
       Else
           'BUSCA LAS PERSONA INVOLUCRADAS EN EL MOVIMIENTO
            sSQL = " SELECT  DISTINCT M.CMOVNRO,  " _
                    & "         CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                    & "                 WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                    & "         P.cPersCod, " _
                    & "         Convert(Char(40),P.CPERSNOMBRE) as cPersNombre, " _
                    & "         MOI.CPersCod " _
                    & " FROM    MOV M JOIN MOVCTA MC    ON M.NMOVNRO = MC.NMOVNRO " _
                    & "         JOIN MOVGasto MOI         ON MOI.NMOVNRO = M.NMOVNRO " _
                    & "         JOIN PERSONA P          ON P.CPERSCOD = MOI.CPERSCOD " _
                    & " WHERE   M.cMovNro = '" & lsMovNro & "' "
              
            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = ""
            Do While Not rs.EOF
                If rs!cDH = "D" Then
                    sDestino = sDestino & Space(8) & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & PstaNombre(rs!cpersNombre)
                Else
                    sOrigen = sOrigen & Space(8) & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & PstaNombre(rs!cpersNombre)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & Space(8) & ImpreFormat("Persona(s) Origen  ", 25) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & Space(8) & ImpreFormat("Persona(s) Destino ", 25) & " :" & sDestino & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
            sSQL = " SELECT  DISTINCT M.CMOVNRO, " _
                & "         CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                 WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "         P.cPersCod , CI.cCtaIFCod , " _
                & "         Convert(Char(40),CTPOI.cConsDescripcion ) as Entidad, " _
                & "         Convert(Char(40),P.CPERSNOMBRE) as cPersNombre, " _
                & "         Convert(char(50),C.cConsDescripcion + ' ' + CI.cCtaIFDesc) as DescCtaIF, " _
                & "         MO.COBJETOCOD " _
                & " FROM    MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "         JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "         JOIN MOVOBJIF MOI       ON MOI.NMOVNRO = MO.NMOVNRO AND MOI.NMOVITEM = MO.NMOVITEM " _
                & "         JOIN CTAIF CI           ON CI.CCTAIFCOD = MOI.CCTAIFCOD AND MOI.CPERSCOD = CI.CPERSCOD AND MOI.CIFTPO = CI.CIFTPO  " _
                & "         JOIN CONSTANTE C        ON C.nConsValor = SUBSTRING(CI.CCTAIFCOD,2,1) AND C.nCONSCOD LIKE '" & gCGTipoCtaIF & "' " _
                & "         JOIN INSTITUCIONFINANC I    ON I.CPERSCOD = CI.CPERSCOD AND I.CIFTPO = CI.CIFTPO " _
                & "         JOIN PERSONA P          ON P.CPERSCOD = I.CPERSCOD " _
                & "         JOIN CONSTANTE CTPOI        ON CTPOI.nConsValor = I.cIFTpo AND CTPOI.nCONSCOD LIKE '" & gCGTipoIF & "' " _
                & " WHERE M.cMovNro = '" & lsMovNro & "' AND MO.cObjetoCod ='" & Format(ObjEntidadesFinancieras, "00") & "'"
            
            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                If rs!cDH = "D" Then
                    sDestino = sDestino & Space(8) & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cpersNombre) & " " & Trim(rs!DescCtaIF)
                Else
                    sOrigen = sOrigen & Space(8) & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cpersNombre) & " " & Trim(rs!DescCtaIF)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & Space(8) & ImpreFormat("Cuenta(s) Origen", 20) & " :" + PrnSet("C+") + sOrigen + PrnSet("C-") + oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & Space(8) & ImpreFormat("Cuenta(s) Destino", 20) & " :" + PrnSet("C+") + sDestino + PrnSet("C-") + oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
            sSQL = "    SELECT  DISTINCT M.CMOVNRO, " _
                & "             CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                     WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "             MOA.cAreaCod + MOA.CAGECOD AS cAreaCod, A.cAreaDescripcion + ' ' + ISNULL(AG.CAGEDESCRIPCION,'' )  as cAreaDescripcion , MO.COBJETOCOD, O.cObjetoDesc " _
                & "     FROM    MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "             JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "             JOIN MovObjAreaAgencia MOA         ON MOA.NMOVNRO = MO.NMOVNRO AND MOA.NMOVITEM = MO.NMOVITEM " _
                & "             JOIN AREAS A            ON A.cAreaCod = MOA.cAreaCod LEFT JOIN AGENCIAS AG ON AG.CAGECOD = MOA.CAGECOD " _
                & "             JOIN OBJETO O ON O.cObjetoCod  = MO.cObjetoCod " _
                & "     WHERE   M.cMovNro = '" & lsMovNro & "' AND Mo.cObjetoCod IN ('" & Format(ObjCMACAgenciaArea, "00") & "','" & Format(ObjCMACArea, "00") & "')"

            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                lsDesObjeto = Trim(rs!cObjetoDesc)
                If rs!cDH = "D" Then
                    sDestino = sDestino & Space(8) & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!cAreaDescripcion)
                Else
                    sOrigen = sOrigen & Space(8) & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!cAreaDescripcion)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & Space(8) & ImpreFormat(lsDesObjeto & " Origen ", 20) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & Space(8) & ImpreFormat(lsDesObjeto & " Destino ", 20) & " :" & sDestino & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
            sSQL = " SELECT    DISTINCT M.CMOVNRO, " _
                & "             CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                     WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "             MO.COBJETOCOD, O.cObjetoDesc " _
                & "   FROM      MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "             JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "             JOIN OBJETO O ON O.cObjetoCod  = MO.cObjetoCod " _
                & "   WHERE     M.cMovNro='" & lsMovNro & "' " _
                & "             AND MO.cObjetoCod NOT IN ('" & Format(ObjCMACAgencias, "00") & "','" & Format(ObjCMACAgenciaArea, "00") & "','" & Format(ObjDescomEfectivo, "00") & "','" & Format(ObjEntidadesFinancieras, "00") & "','" & Format(ObjBienesServicios, "00") & "','" & Format(ObjPersona, "00") & "')"
            
            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                lsDesObjeto = "OBJETOS"
                If rs!cDH = "D" Then
                    sDestino = sDestino & Space(8) & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cObjetoCod) & "-" & Trim(rs!cObjetoDesc)
                Else
                    sOrigen = sOrigen & Space(8) & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cObjetoCod) & "-" & Trim(rs!cObjetoDesc)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Origen ", 20) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Destino", 20) & " :" & sDestino & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
        End If ' del pscab1
    
       'Tipo de Cambio
        sSQL = "SELECT nMovTpoCambio, m.cOpeCod FROM MovTpoCambio MC JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE M.cMovNro = '" & lsMovNro & "'"
        Set rs = oConect.CargaRecordSet(sSQL)
        If Not rs.EOF Then
            If rs!cOpeCod = gContAjusteTipoCambio Then
            Else
               sAsiento = sAsiento & Space(8) & "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "      T.C.Mercado: " & Format(rs!nMovTpoCambio, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
            End If
            nLi = nLi + 1
        Else
            If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Then
                sAsiento = sAsiento & Space(8) & "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "     T.C.Compra : " & Format(lnCambioCompra, "##,###,##0.000") & "     T.C.Venta : " & Format(lnCambioVenta, "##,###,##0.000") & "     T.C.C SBS : " & Format(lnCambioCompraSBS, "##,###,##0.000") & "     T.C.V SBS : " & Format(lnCambioVentaSBS, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
        End If
       nLi = nLi + 1
       'Datos de Documentos
       sSQL = "SELECT b.cDocAbrev, a.cDocNro, a.dDocFecha,b.nDocTpo FROM MovDoc a join mov m on m.nmovnro =a.nmovnro  , " & vsBaseComunes & "Documento b " _
            & "WHERE m.cMovNro = '" & lsMovNro & "' and b.nDocTpo = a.nDocTpo "
       
        Set rsDoc = oConect.CargaRecordSet(sSQL)
        If Not rsDoc.EOF Then
          sDoc = ""
          nLi = 1
          Do While Not rsDoc.EOF
             sDoc = sDoc & rsDoc!cDocAbrev & "_" & rsDoc!cDocNro & " " & rsDoc!dDocFecha & "  "
             If nLi Mod 4 = 0 Then
                sDoc = sDoc + oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
             End If
            '---  Sacar Numero de Orden de Pago  JEOM
            If rsDoc!nDocTpo = 48 Then
                Dim nOP As String
                nOP = rsDoc!cDocNro
            End If
            '------------------------------------------
             rsDoc.MoveNext
          Loop
                          
          sAsiento = sAsiento & Space(8) & ImpreFormat("Documentos", 18) & " :" + sDoc + oImpresora.gPrnSaltoLinea
          nLi = nLi + 1
        End If

      Dim NroMov As Long
      Dim oMov As DMov
      Dim sSQL1 As String
      Dim rs1 As ADODB.Recordset
          
      Set oMov = New DMov
      Set rs1 = New ADODB.Recordset
            
      NroMov = oMov.GetnMovNro(psMovNro)
       sSQL1 = "SELECT p.cPersNombre,p.cPersCod FROM MovGasto m inner join persona p on p.cPErscod =m.cPerscod WHERE m.nMovNro = '" & NroMov & "'"
       Set rs1 = oConect.CargaRecordSet(sSQL1)
       If rs1.EOF Then
          Exit Function
       End If
            
            
            Dim sMedio As String
            sMedio = sMedio & oImpresora.gPrnSaltoLinea & Space(73) & nOP & Space(62) & "162398" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            sMedio = sMedio & Space(135) & lsFecha & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            sMedio = sMedio & Space(10) & rs1!cpersNombre & oImpresora.gPrnSaltoLinea
            '''sMedio = sMedio & Space(130) & IIf(psMoneda = 1, "S/. ", "$. ") & Format(pnMonto, "#,##0.00") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            sMedio = sMedio & Space(130) & IIf(psMoneda = 1, gcPEN_SIMBOLO & " ", "$. ") & Format(pnMonto, "#,##0.00") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016

                    
            
            
      '-- **************************************
        rsDoc.Close
        
        sAsiento = sMedio & sAsiento
       'Glosa
       'sAsiento = sAsiento & ImpreGlosa(lsGlosa, pnColPage, ImpreFormat("Glosa       ", 20) & " :", , , nLi) '+ oImpresora.gPrnSaltoLinea
       nLi = nLi + 1
       lsMovBilletaje = ""
       lsMovBilletaje = ImprimeBilletaje(Mid(lsOpeCod, 3, 1), lsMovNro, pnColPage, , nLi)
       If lsMovBilletaje <> "" Then
            sAsiento = sAsiento + lsMovBilletaje
       End If
        'MovCta para MN
    sSQL = "SELECT mc.nMovItem, mc.cCtaContCod, ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),'') cCtaContDesc, mc.nMovImporte " _
         & "FROM   MovCta mc JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE  M.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
    Set rs = oConect.CargaRecordSet(sSQL)
    If rs.EOF Then
       'MsgBox "No existen datos de Movimiento ", vbInformation, "Error"
       Exit Function
    End If
    nTot = 0
    nTotH = 0
    sAsiento = sAsiento '+ sCabecera
    nLi = nLi + 4
    sAsiento = sAsiento & Space(9) & " EN MONEDA NACIONAL " & oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento & Space(9) & "--------------------" & oImpresora.gPrnSaltoLinea
    'sAsiento = sAsiento & oImpresora.gPrnSaltoLinea
    Do While Not rs.EOF
       DoEvents
       If nLi > pnLinPage - 6 Then
           nLi = 5
           sAsiento = sAsiento + oImpresora.gPrnSaltoPagina & sCabecera
       End If
       sAsiento = sAsiento & Space(9) & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " & Justifica(rs!cCtaContDesc, 71) & " " _
            & Right(Space(16) & IIf(rs!nMovImporte > 0, Format(rs!nMovImporte, "##,###,##0.00"), ""), 16) & " " _
            & Right(Space(16) & IIf(rs!nMovImporte < 0, Format(rs!nMovImporte * -1, "##,###,##0.00"), ""), 16) _
            & oImpresora.gPrnSaltoLinea
       If rs!nMovImporte > 0 Then
          nTot = nTot + Val(rs!nMovImporte)
       Else
          nTotH = nTotH + Val(rs!nMovImporte) * -1
       End If
       nLi = nLi + 1
       rs.MoveNext
    Loop
    sAsiento = sAsiento & Space(9) & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento & Space(9) & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & oImpresora.gPrnSaltoLinea
    nLi = nLi + 2

    sSQL = "SELECT mc.nMovItem, mc.cCtaContCod, LTRIM(RTRIM(ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),''))) cCtaContDesc, me.nmovMEimporte " _
         & "FROM   mov m join MovCta mc on mc.nmovnro=m.nmovnro JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
         & "WHERE  m.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
    Set rs = oConect.CargaRecordSet(sSQL)
    sTexto = ""
    nTot = 0
    nTotH = 0
    If Not rs.EOF Then
       sAsiento = sAsiento & Space(9) & " EN MONEDA EXTRANJERA " & oImpresora.gPrnSaltoLinea
       sAsiento = sAsiento & Space(9) & "----------------------" & oImpresora.gPrnSaltoLinea
       nLi = nLi + 2
    End If
    Do While Not rs.EOF
       DoEvents
       If nLi > pnLinPage - 6 Then
          nLi = 5
          sTexto = sTexto + oImpresora.gPrnSaltoPagina
          sTexto = sTexto + sCabecera
       End If
       sTexto = sTexto & Space(9) & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " _
              & Justifica(rs!cCtaContDesc, 71) & " " _
              & IIf(rs!nMovMEImporte > 0, PrnVal(rs!nMovMEImporte, 16, 2), Space(16)) & " " _
              & IIf(rs!nMovMEImporte < 0, PrnVal(rs!nMovMEImporte * -1, 16, 2), Space(16)) _
              & oImpresora.gPrnSaltoLinea
       If rs!nMovMEImporte > 0 Then
          nTot = nTot + Val(rs!nMovMEImporte)
       Else
          nTotH = nTotH + Val(rs!nMovMEImporte) * -1
       End If
       nLi = nLi + 1
       rs.MoveNext
    Loop
    If nTot > 0 Or nTotH > 0 Then
       sAsiento = sAsiento & sTexto
       sAsiento = sAsiento & Space(9) & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
       sAsiento = sAsiento & Space(9) & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) + oImpresora.gPrnSaltoLinea
       nLi = nLi + 3
    End If
    If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
    
    If lsOpeCod <> 421113 Then
        sAsiento = sAsiento & Space(9) & "====================================================================================================================================" & oImpresora.gPrnSaltoLinea
        
        'sAsiento = sAsiento + ImprePiePag(pnColPage, psPiePag, psOtraFirma)
    End If
    sAsiento = sAsiento + "   " + oImpresora.gPrnSaltoLinea

    Set oDGeneral = Nothing
    Set oConect = Nothing
    ImprimeAsientoContableNew = sAsiento
    
    Exit Function
ImprimeAsientoContableErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeAsientoContable Method")
End Function


Public Function ProvContxCtasContables(ByVal psFechaIni As Date, ByVal psFechaFin As Date) As String
    On Error GoTo ProvContxCtasContablesErr
    Dim sSQL As String
    Dim sCabecera As String
    Dim sAsiento As String
    Dim nLi As Long
    Dim lsFechaIni As String
    Dim lsFechaFin As String
    
    Dim rs       As New ADODB.Recordset
    Dim oConect As DConecta
    Set oConect = New DConecta
    
    If oConect.AbreConexion = False Then Exit Function
    
    sCabecera = ""
    sAsiento = ""
    nLi = 0
    
    lsFechaIni = Format(psFechaIni, "yyyy/mm/dd")
    lsFechaFin = Format(psFechaFin, "yyyy/mm/dd")
       
      
     sSQL = " Select  Sum(nDebe-nHaber)Total,cCtaCnt Cuenta from AsientoDN" & _
            " where dFecha>= '" & lsFechaIni & "' and dFecha <='" & lsFechaFin & "' and cctaCnt like '2118%' and cTipo=0 " & _
            " Group by cCtacnt"
        
       Set rs = oConect.CargaRecordSet(sSQL)
       If rs.EOF Then
          Exit Function
       End If
      
    sCabecera = sCabecera & Space(5) & "DEBE - HABER" & Space(10) & "Cta. Contable" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera & String(35, "_") & oImpresora.gPrnSaltoLinea
    Dim Total As String * 10
    Dim Cuenta As String * 15
    Dim Final As String
   

    Do While Not rs.EOF
       DoEvents
       If nLi > 60 Then
           nLi = 5
           sAsiento = sCabecera & sAsiento & oImpresora.gPrnSaltoPagina
       End If
       Total = rs!Total
       Cuenta = rs!Cuenta
       sAsiento = sAsiento & Total & Space(12) & Cuenta & oImpresora.gPrnSaltoLinea

       nLi = nLi + 1
       rs.MoveNext
    Loop
       
    sAsiento = sCabecera & sAsiento
    ProvContxCtasContables = sAsiento
    Exit Function
ProvContxCtasContablesErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ProvContxCtasContables Method")
End Function

Public Function ProvContxOperacion(ByVal psFechaIni As Date, ByVal psFechaFin As Date) As String
    On Error GoTo ProvContxOperacionErr
    Dim sSQL As String
    Dim sCabecera As String
    Dim sAsiento As String
    Dim nLi As Long
    Dim lsFechaIni As String
    Dim lsFechaFin As String
    
    Dim rs       As New ADODB.Recordset
    Dim oConect As DConecta
    Set oConect = New DConecta
    
    If oConect.AbreConexion = False Then Exit Function
    
    sCabecera = ""
    sAsiento = ""
    nLi = 0
    
    lsFechaIni = Format(psFechaIni, "yyyymmdd")
    lsFechaFin = Format(psFechaFin, "yyyymmdd")
       
      
     sSQL = " Select Sum(nDebe) Debe,Sum(nHaber)Haber,cOpeCod Operacion" & _
            " From AsientoDN " & _
            " Where dFecha >='" & lsFechaIni & "' and dFecha <= '" & lsFechaFin & "' and cCtaCnt like '2118%10' and cTipo=0 " & _
            " Group by cOpeCod"
        
       Set rs = oConect.CargaRecordSet(sSQL)
       If rs.EOF Then
          Exit Function
       End If
      
    sCabecera = sCabecera & Space(5) & "DEBE" & Space(13) & "HABER" & Space(12) & "OPERACION" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera & String(50, "_") & oImpresora.gPrnSaltoLinea
    
    Dim Debe As String * 10
    Dim Haber As String * 10
    Dim Cuenta As String * 15
    Dim Operacion As String * 15
    
    Do While Not rs.EOF
       DoEvents
       If nLi > 60 Then
           nLi = 5
           sAsiento = sCabecera & sAsiento & oImpresora.gPrnSaltoPagina
       End If
       Debe = rs!Debe
       Haber = rs!Haber
       Operacion = rs!Operacion
       sAsiento = sAsiento & Debe & Space(7) & Haber & Space(7) & Operacion & oImpresora.gPrnSaltoLinea
          
       nLi = nLi + 1
       rs.MoveNext
    Loop
    
    sAsiento = sCabecera & sAsiento
    ProvContxOperacion = sAsiento
    Exit Function
ProvContxOperacionErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ProvContxCtasContables Method")
End Function

Public Function InteresDevvengadoPFxFxAG(ByVal psFechaIni As Date, ByVal psFechaFin As Date) As String
    On Error GoTo InteresDevvengadoPFxFxAGErr
    Dim sSQL As String
    Dim sCabecera As String
    Dim sAsiento As String
    Dim nLi As Long
    Dim lsFechaIni As String
    Dim lsFechaFin As String
    
    Dim rs       As New ADODB.Recordset
    Dim oConect As DConecta
    Set oConect = New DConecta
    
    If oConect.AbreConexion = False Then Exit Function
    
    sCabecera = ""
    sAsiento = ""
    nLi = 0
    
    lsFechaIni = Format(psFechaIni, "yyyymmdd")
    lsFechaFin = Format(psFechaFin, "yyyymmdd")
       
      
     sSQL = " Select dEstad a, cCodAge ag, SUM(nIntDev) Int " & _
            " From CapEstadSaldo " & _
            " Where nMoneda = 1 And nProducto = 233" & _
            " and dEstad >= '" & lsFechaIni & "' and dEstad <= '" & lsFechaFin & "' " & _
            " group by dEstad,ccodAge" & _
            " Order by a,ag"
        
       Set rs = oConect.CargaRecordSet(sSQL)
       If rs.EOF Then
          Exit Function
       End If
      
    sCabecera = sCabecera & "FECHA" & Space(13) & "AGENCIA" & Space(12) & "MONTO" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera & String(50, "_") & oImpresora.gPrnSaltoLinea
    Dim Fecha As String * 10
    Dim Agencia As String * 10
    Dim Monto As String * 10
    
    
    Do While Not rs.EOF
       DoEvents
       If nLi > 60 Then
           nLi = 5
           sAsiento = sAsiento & oImpresora.gPrnSaltoPagina & sCabecera
       End If
       Fecha = rs!a
       Agencia = rs!ag
       Monto = rs!Int
       sAsiento = sAsiento & Fecha & Space(7) & Agencia & Space(7) & Monto & oImpresora.gPrnSaltoLinea
          
       nLi = nLi + 1
       rs.MoveNext
    Loop
    
    sAsiento = sCabecera & sAsiento
    InteresDevvengadoPFxFxAG = sAsiento
    Exit Function
InteresDevvengadoPFxFxAGErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:InteresDevvengadoPFxFxAG Method")
End Function


Public Function InteresCTSxFxAG(ByVal psFechaIni As Date, ByVal psFechaFin As Date) As String
    On Error GoTo InteresCTSxFxAGErr
    Dim sSQL As String
    Dim sCabecera As String
    Dim sAsiento As String
    Dim nLi As Long
    Dim lsFechaIni As String
    Dim lsFechaFin As String
    
    Dim rs       As New ADODB.Recordset
    Dim oConect As DConecta
    Set oConect = New DConecta
    
    If oConect.AbreConexion = False Then Exit Function
    
    sCabecera = ""
    sAsiento = ""
    nLi = 0
    
    lsFechaIni = Format(psFechaIni, "yyyymmdd")
    lsFechaFin = Format(psFechaFin, "yyyymmdd")
       
      
     sSQL = " Select dEstad a, cCodAge ag, SUM(nIntDev) Int" & _
            " From CapEstadsaldo " & _
            " where nMoneda=1 and nProducto =234 and dEstad >='" & lsFechaIni & "' and dEstad < '" & lsFechaFin & "'" & _
            " Group by dEstad, cCodAge" & _
            " Order by a,ag"
        
       Set rs = oConect.CargaRecordSet(sSQL)
       If rs.EOF Then
          Exit Function
       End If
      
    sCabecera = sCabecera & "FECHA" & Space(13) & "AGENCIA" & Space(12) & "MONTO" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera & String(50, "_") & oImpresora.gPrnSaltoLinea
    
    Dim Fecha As String * 10
    Dim Agencia As String * 10
    Dim Monto As String * 10
    
    
    Do While Not rs.EOF
       DoEvents
       If nLi > 60 Then
           nLi = 5
           sAsiento = sAsiento & oImpresora.gPrnSaltoPagina & sCabecera
       End If
       Fecha = rs!a
       Agencia = rs!ag
       Monto = rs!Int
       sAsiento = sAsiento & Fecha & Space(7) & Agencia & Space(7) & Monto & oImpresora.gPrnSaltoLinea
          
       nLi = nLi + 1
       rs.MoveNext
    Loop
    
    sAsiento = sCabecera & sAsiento
    InteresCTSxFxAG = sAsiento
    Exit Function
InteresCTSxFxAGErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:InteresCTSxFxAG Method")
End Function
'By Capi 04022008
Public Function ImprimePlazoFijoIntCash(ByVal psMoneda As String, ByVal psNomCmac As String, _
                                         ByVal psNomAge As String, ByVal pdFecsis As Date, _
                                         ByVal psCodUser As String)



  
    Dim lsFecha As String
    Dim lsMoneda As String
    Dim R As New ADODB.Recordset
    Dim oDCap As DCapMantenimiento
    Dim vAgencia As String
    Dim sCadImp As String
    Dim j As Integer
    Dim K As Integer
    Dim lnTotSaldo As Double
    Dim lnTotcont As Double
    Dim lnTotPagado As Double
    Dim lnTotDiferido As Double
    Dim lnTotaDiferir As Double
    
    
    Set oDCap = New DCapMantenimiento
    Set R = oDCap.ObtenerPlazoFijoIntCash("1", psMoneda)
    Set oDCap = Nothing
    lsFecha = "Periodo: " & Trim(Str(R!Mes)) & "/" & Trim(Str(R!Anno)) 'Determinado a que fecha se emitira el reporte.
    lsMoneda = IIf(psMoneda = "1", "Moneda Nacional", "Moneda Extranjera")
        
    
    Set oDCap = New DCapMantenimiento
    Set R = oDCap.ObtenerPlazoFijoIntCash("2", psMoneda)
    Set oDCap = Nothing
    
    
    
    sCadImp = ""
    sCadImp = sCadImp & gPrnCondensadaON
    
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecsis, psCodUser, "REPORTE DE PLAZOS FIJOS VIGENTES - INTERES CASH", psNomCmac, , True, , lsFecha & " En: " & lsMoneda)
        sCadImp = sCadImp & String(166, "-") & oImpresora.gPrnSaltoLinea
        '                    0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                              1         2         3         4         5         6         7         8         9         10        11        12        13        14        15        16        17
        sCadImp = sCadImp & "NRO    CUENTA              CLIENTE                        APERTURA      PLAZO   PROX CANCEL      M.APERTURA    S.CONTABLE    INT PAGADO   INT DIFERIDO    A DIFERIR   " & oImpresora.gPrnSaltoLinea
        sCadImp = sCadImp & String(166, "-") & oImpresora.gPrnSaltoLinea
    

    K = 0
    Do While Not R.EOF
   
        
        sCadImp = sCadImp & "Agencia  : " & R!Agencia & " " & R!AgenciaDesc & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        vAgencia = R!Agencia
        j = 0
        
        lnTotSaldo = 0
        lnTotPagado = 0
        lnTotDiferido = 0
        lnTotaDiferir = 0
        
        Do While R!Agencia = vAgencia
            j = j + 1
            K = K + 1
            If K = 51 Then
                K = 0
                sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
                Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecsis, psCodUser, "REPORTE DE PLAZOS FIJOS VIGENTES - INTERES CASH", psNomCmac, , True, , "HASTA:" & lsFecha & " " & lsMoneda)
                sCadImp = sCadImp & String(166, "-") & oImpresora.gPrnSaltoLinea
                sCadImp = sCadImp & "NRO    CUENTA              CLIENTE                        APERTURA      PLAZO   PROX CANCEL      M.APERTURA    S.CONTABLE    INT PAGADO   INT DIFERIDO    A DIFERIR   " & oImpresora.gPrnSaltoLinea
                sCadImp = sCadImp & String(166, "-") & oImpresora.gPrnSaltoLinea
            End If
            sCadImp = sCadImp & ImpreFormat(j, 4, 0)
            sCadImp = sCadImp & ImpreFormat(R!Cuenta, 18)
            sCadImp = sCadImp & ImpreFormat(R!Cliente, 30)
            sCadImp = sCadImp & ImpreFormat(Format(R!Apertura, "dd/mm/yyyy"), 10)
            sCadImp = sCadImp & ImpreFormat(R!Plazo, 10, 0)
            sCadImp = sCadImp & ImpreFormat(Format(R!dCancPF, "dd/mm/yyyy"), 10)
            sCadImp = sCadImp & ImpreFormat(R!mApertura, 12, 2)
            sCadImp = sCadImp & ImpreFormat(R!scontable, 12, 2)
            sCadImp = sCadImp & ImpreFormat(R!Pagado, 12, 2)
            sCadImp = sCadImp & ImpreFormat(R!Diferido, 12, 2)
            sCadImp = sCadImp & ImpreFormat(R!Saldo, 12, 2)
            sCadImp = sCadImp & oImpresora.gPrnSaltoLinea
            
            lnTotSaldo = lnTotSaldo + R!mApertura
            lnTotcont = lnTotcont + R!scontable
            lnTotPagado = lnTotPagado + R!Pagado
            lnTotDiferido = lnTotDiferido + R!Diferido
            lnTotaDiferir = lnTotaDiferir + R!Saldo
            
            R.MoveNext
            
            If R.EOF Then
                Exit Do
            End If
        Loop
        sCadImp = sCadImp & String(166, "-") & oImpresora.gPrnSaltoLinea
        sCadImp = sCadImp & Space(92)
        
        sCadImp = sCadImp & ImpreFormat(Format(lnTotSaldo, "#,###,###.##"), 12, 2)
        sCadImp = sCadImp & Space(2)
        sCadImp = sCadImp & ImpreFormat(Format(lnTotcont, "#,###,###.##"), 12, 2)
        sCadImp = sCadImp & Space(2)
        sCadImp = sCadImp & ImpreFormat(Format(lnTotPagado, "#,###,###.##"), 12, 2)
        sCadImp = sCadImp & Space(2)
        sCadImp = sCadImp & ImpreFormat(Format(lnTotDiferido, "#,###,###.##"), 12, 2)
        sCadImp = sCadImp & Space(2)
        sCadImp = sCadImp & ImpreFormat(Format(lnTotaDiferir, "#,###,###.##"), 12, 2)
        
        sCadImp = sCadImp & oImpresora.gPrnSaltoLinea
        K = K + 4
    Loop
    
    R.Close
    Set R = Nothing
    sCadImp = sCadImp & gPrnCondensadaOFF
    ImprimePlazoFijoIntCash = sCadImp
End Function
'By Capi 04022008
Private Sub ImprimeCabeceraDocumento(ByRef psCadImp As String, ByVal psNomAge As String, _
                                    ByVal psFechaHora As String, ByVal psCodUsu, ByVal psTitulo As String, ByVal psNomCmac As String, _
                                    Optional psTab As String = "", Optional pbCondensado As Boolean = True, Optional psCodRepo As String = "", _
                                    Optional psSubTitulo As String = "") 'DAOR 20070816, Se aumentó subtitulo
    Dim EspHorxPag As Integer
    EspHorxPag = 170
    psCadImp = psCadImp & oImpresora.gPrnSaltoLinea
    If Len(psCodRepo) > 0 Then
        psTitulo = psCodRepo & " - " & psTitulo
    End If
    If pbCondensado Then
        psCadImp = psCadImp & psTab & psNomCmac & Space(70) & "FECHA :" & psFechaHora & Chr$(10)
        psCadImp = psCadImp & psTab & ImpreFormat(psNomAge, 45, 0) & Space(51) & "USUARIO : " & psCodUsu & Chr$(10)
        psCadImp = psCadImp & psTab & Space((EspHorxPag - Len(psTitulo)) / 2 - 18) & psTitulo & Chr$(10)
        psCadImp = psCadImp & psTab & Space((EspHorxPag - Len(psSubTitulo)) / 2 - 18) & psSubTitulo & Chr$(10) 'DAOR 20070816
    Else
        psCadImp = psCadImp & psTab & psNomCmac & Space(76) & "FECHA   :" & psFechaHora & Chr$(10)
        psCadImp = psCadImp & psTab & ImpreFormat(psNomAge, 40, 0) & Space(36) & "USUARIO : " & psCodUsu & Chr$(10)
        psCadImp = psCadImp & psTab & Space((120 - Len(psTitulo)) / 2 - 18) & psTitulo & Chr$(10)
        psCadImp = psCadImp & psTab & Space((120 - Len(psSubTitulo)) / 2 - 18) & psSubTitulo & Chr$(10) 'DAOR 20070816
    End If
End Sub
'By Capi 04022008
Private Sub ImprimeCabeceraDocumentoAncho(ByRef psCadImp As String, ByVal psNomAge As String, _
                                    ByVal psFechaHora As String, ByVal psCodUsu, ByVal psTitulo As String, ByVal psNomCmac As String, _
                                    Optional psTab As String = "", Optional pbCondensado As Boolean = True, Optional psCodRepo As String = "", _
                                    Optional psSubTitulo As String = "") 'DAOR 20070816, Se aumentó subtitulo
    Dim EspHorxPag As Integer
    EspHorxPag = 280
    'psCadImp = psCadImp & oImpresora.gPrnSaltoLinea
    If Len(psCodRepo) > 0 Then
        psTitulo = psCodRepo & " - " & psTitulo
    End If
    If pbCondensado Then
        psCadImp = psCadImp & psTab & psNomCmac & Space(190) & "FECHA :" & psFechaHora & oImpresora.gPrnSaltoLinea
        psCadImp = psCadImp & psTab & ImpreFormat(psNomAge, 45, 0) & Space(169) & "USUARIO : " & psCodUsu & oImpresora.gPrnSaltoLinea
        psCadImp = psCadImp & psTab & Space((EspHorxPag - Len(psTitulo)) / 2 - 18) & psTitulo & oImpresora.gPrnSaltoLinea
        psCadImp = psCadImp & psTab & Space((EspHorxPag - Len(psSubTitulo)) / 2 - 18) & psSubTitulo & oImpresora.gPrnSaltoLinea 'DAOR 20070816
    Else
        psCadImp = psCadImp & psTab & psNomCmac & Space(76) & "FECHA   :" & psFechaHora & oImpresora.gPrnSaltoLinea
        psCadImp = psCadImp & psTab & ImpreFormat(psNomAge, 40, 0) & Space(36) & "USUARIO : " & psCodUsu & oImpresora.gPrnSaltoLinea
        psCadImp = psCadImp & psTab & Space((120 - Len(psTitulo)) / 2 - 18) & psTitulo & oImpresora.gPrnSaltoLinea
        psCadImp = psCadImp & psTab & Space((120 - Len(psSubTitulo)) / 2 - 18) & psSubTitulo & oImpresora.gPrnSaltoLinea 'DAOR 20070816
    End If
End Sub
'By Capi 08022008
Private Sub ImprimeTotalesCarteraDetxIntDSP(psCadImp As String, ByVal pnTotAprobado As Double, _
                                    ByVal pnTotSaldoK As Double, ByVal pnTotDevengado As Double, _
                                    ByVal pnTotSuspenso As Double, ByVal pnTotProvision As Double, ByVal pnCuantos As Double, ByVal psGrupo As String)

        psCadImp = psCadImp & Space(79)
        psCadImp = psCadImp & String(90, "-") & oImpresora.gPrnSaltoLinea
        psCadImp = psCadImp & Space(79)
        psCadImp = psCadImp & " MONTO APRO.           SALDO K       DEVENGADO      SUSPENSO         PROVISION " & oImpresora.gPrnSaltoLinea
        psCadImp = psCadImp & Space(79)
        psCadImp = psCadImp & String(90, "-") & oImpresora.gPrnSaltoLinea
        psCadImp = psCadImp & psGrupo & pnCuantos & " Caso(s)" & oImpresora.gPrnSaltoLinea
        psCadImp = psCadImp & Space(82)
        
        psCadImp = psCadImp & ImpreFormat(Format(pnTotAprobado, "#,###,###.##"), 12, 2)
        psCadImp = psCadImp & Space(3)
        psCadImp = psCadImp & ImpreFormat(Format(pnTotSaldoK, "#,###,###.##"), 12, 2)
        psCadImp = psCadImp & Space(3)
        psCadImp = psCadImp & ImpreFormat(Format(pnTotDevengado, "#,###,###.##"), 12, 2)
        psCadImp = psCadImp & Space(3)
        psCadImp = psCadImp & ImpreFormat(Format(pnTotSuspenso, "#,###,###.##"), 12, 2)
        psCadImp = psCadImp & Space(3)
        psCadImp = psCadImp & ImpreFormat(Format(pnTotProvision, "#,###,###.##"), 12, 2)
        psCadImp = psCadImp & gPrnSaltoLinea
        
       
        
        'Reiniciando en 0
        pnTotAprobado = 0
        pnTotSaldoK = 0
        pnTotDevengado = 0
        pnTotSuspenso = 0
        pnTotProvision = 0
        
End Sub
'By Capi 08022008
Public Function ImprimeCarteraDetxIntDSP(pnTipCambio As Double, ByVal psNomCmac As String, _
                                         ByVal psNomAge As String, ByVal pdFecsis As Date, _
                                         ByVal psCodUser As String, _
                                         Optional psExcel As Boolean = False)
  
    Dim lsFecha As String
    Dim lnPeriodo As Integer
     'By capi 17112008
    Dim lnMes As Integer
    Dim lnAnno As Integer
    '
    Dim lsMoneda As String
    Dim R As New ADODB.Recordset
    Dim oDCre As DCreditos
    Dim vAgencia As String
    Dim vEstado As String
    Dim vMoneda As String
    Dim vCadena As String
    
    Dim vTipo As String
    Dim vSituacion As String
    Dim sCadImp As String
    Dim lsCadImp As String
    
    Dim j As Double
    Dim K As Integer
    Dim lnCuantos As Integer
    Dim lnGCuantos As Double
    Dim sw As Integer
    
    Dim lnTotAAprobado As Double
    Dim lnTotASaldoK As Double
    Dim lnTotADevengado As Double
    Dim lnTotASuspenso As Double
    Dim lnTotAProvision As Double
    
    Dim lnTotGAprobado As Double
    Dim lnTotGSaldoK As Double
    Dim lnTotGDevengado As Double
    Dim lnTotGSuspenso As Double
    Dim lnTotGProvision As Double
    
      
    'By capi 17112008
    lnMes = Month(pdFecsis)
    lnAnno = Year(pdFecsis)
    '
    
        
    
    Set oDCre = New DCreditos
    Set R = oDCre.ObtenerCalCarteraxIntDevSupProv("1", 0, 0, lnMes, lnAnno)
    Set oDCre = Nothing
    lsMoneda = "Expresado en Moneda Nacional al T.C. : "
        
        
    lsFecha = "Periodo: " & lnMes & "/" & lnAnno
   
  
    
    
    If R!Cuantos = 0 Then
        MsgBox "Cierre de Calificacion Pendiente - No Existe Registros a Procesar", vbInformation, "Aviso"
        R.Close: Set R = Nothing
        Exit Function
    Else
        MsgBox "Se Verifico Cierre Calificacion - Registros a Procesar >>>>>> " & R!Cuantos, vbInformation, "Aviso"
    End If

        
        
    
    Set oDCre = New DCreditos
    Set R = oDCre.ObtenerCalCarteraxIntDevSupProv("2", 0, pnTipCambio, lnMes, lnAnno)
    Set oDCre = Nothing
    
    If psExcel = True Then
        Dim lMatCabecera As Variant
        ReDim lMatCabecera(21, 2)
        lMatCabecera(0, 0) = "Agencia": lMatCabecera(0, 1) = ""
        lMatCabecera(1, 0) = "Estado": lMatCabecera(1, 1) = ""
        lMatCabecera(2, 0) = "Moneda": lMatCabecera(2, 1) = ""
        lMatCabecera(3, 0) = "Tipo": lMatCabecera(3, 1) = ""
        lMatCabecera(4, 0) = "Condicion": lMatCabecera(4, 1) = ""
        lMatCabecera(5, 0) = "Cuenta": lMatCabecera(5, 1) = ""
        lMatCabecera(6, 0) = "Cliente": lMatCabecera(6, 1) = ""
        lMatCabecera(7, 0) = "Codigo": lMatCabecera(7, 1) = ""
        lMatCabecera(8, 0) = "Calificacion": lMatCabecera(8, 1) = ""
        lMatCabecera(9, 0) = "Vigencia": lMatCabecera(9, 1) = ""
        lMatCabecera(10, 0) = "Vencimiento": lMatCabecera(10, 1) = ""
        lMatCabecera(11, 0) = "Pagado": lMatCabecera(11, 1) = ""
        lMatCabecera(12, 0) = "Proxima": lMatCabecera(12, 1) = ""
        lMatCabecera(13, 0) = "Mora": lMatCabecera(13, 1) = ""
        lMatCabecera(14, 0) = "Tasa": lMatCabecera(14, 1) = ""
        lMatCabecera(15, 0) = "Monto_Aprobado": lMatCabecera(15, 1) = ""
        lMatCabecera(16, 0) = "Saldo_Capital": lMatCabecera(16, 1) = ""
        lMatCabecera(17, 0) = "Int_Devengado": lMatCabecera(17, 1) = ""
        lMatCabecera(18, 0) = "Int_Suspenso": lMatCabecera(18, 1) = ""
        lMatCabecera(19, 0) = "Provision": lMatCabecera(19, 1) = ""
        lMatCabecera(20, 0) = "Tasa_Prov": lMatCabecera(20, 1) = ""
        
        Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, psCodUser, gdFecSis, "LISTADO CALIFICACION CARTERA: INTERES DEVG, SUSP y PROVISIONES - " & lsFecha & lsMoneda & pnTipCambio, "", "DetalleCarteraIntProv", lMatCabecera, R, 2, True, , True)
        Exit Function
        
    End If
       
    sCadImp = ""
    lsCadImp = ""
    
        

    
    j = 0
    sw = 0
    K = 0
    lnCuantos = 0
    lnGCuantos = 0
    
     
    lnTotGAprobado = 0
    lnTotGSaldoK = 0
    lnTotGDevengado = 0
    lnTotGSuspenso = 0
    lnTotGProvision = 0
    
    
            
    
    Do While Not R.EOF
        
        vAgencia = R!Agencia
        vEstado = R!Estado
        vMoneda = R!Moneda
        vTipo = R!Tipo
        vSituacion = R!Condicion
       
        vCadena = R!Agencia & R!Estado & R!Moneda & R!Tipo & R!Condicion
        
        
    lnTotAAprobado = 0
    lnTotASaldoK = 0
    lnTotADevengado = 0
    lnTotASuspenso = 0
    lnTotAProvision = 0
   

      
        Do While R!Agencia & R!Estado & R!Moneda & R!Tipo & R!Condicion = vCadena
            j = j + 1
            K = K + 1
            lnCuantos = lnCuantos + 1
            lnGCuantos = lnGCuantos + 1
            If sw = 0 Or sw = 2 Then
                sw = 1
                'sCadImp = sCadImp & oImpresora.gPrnCondensadaON
                Call ImprimeCabeceraDocumentoAncho(sCadImp, psNomAge, pdFecsis, psCodUser, "LISTADO CALIFICACION CARTERA: INTERES DEVG, SUSP y PROVISIONES - " & lsFecha, psNomCmac, , True, , lsMoneda & pnTipCambio)
                
                sCadImp = sCadImp & String(255, "-") & oImpresora.gPrnSaltoLinea
                '                    0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
                '                              1         2         3         4         5         6         7         8         9         10        11        12        13        14        15        16        17
                sCadImp = sCadImp & "   CREDITO             NOMBRE o RAZON SOCIAL CLIENTE           COD. CLIENTE        MONTO APRO.     SALDO K   CALIFICA.         F.VIGEN    VENC CRED   PAGADO       PROX. CUOTA     MORA   TASA INT.  DEVENGADO       SUSPENSO        PROVISION   TASA PROV." & oImpresora.gPrnSaltoLinea
                sCadImp = sCadImp & String(255, "-") & oImpresora.gPrnSaltoLinea
                
                
                sCadImp = sCadImp & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                sCadImp = sCadImp & "Agencia      : " & R!Agencia & oImpresora.gPrnSaltoLinea
                sCadImp = sCadImp & ">Estado      : " & R!Estado & oImpresora.gPrnSaltoLinea
                If R!Moneda = "1" Then
                    sCadImp = sCadImp & ">>           : Moneda Nacional" & oImpresora.gPrnSaltoLinea
                Else
                    sCadImp = sCadImp & ">>           : Moneda Extranjera" & oImpresora.gPrnSaltoLinea
                End If
                sCadImp = sCadImp & ">>>Tipo Cred : " & R!Tipo & oImpresora.gPrnSaltoLinea
                sCadImp = sCadImp & ">>>>Situacion: " & R!Condicion & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            End If
            
                
            'sCadImp = sCadImp & ImpreFormat(J, 4, 0)
            
            
            sCadImp = sCadImp & ImpreFormat(R!Cuenta, 18)
            sCadImp = sCadImp & ImpreFormat(R!Cliente, 40)
            sCadImp = sCadImp & ImpreFormat(R!Codigo, 13)
            sCadImp = sCadImp & ImpreFormat(R!Monto_Aprobado, 12, 2)
            sCadImp = sCadImp & ImpreFormat(R!Saldo_Capital, 12, 2)
            sCadImp = sCadImp & ImpreFormat(R!Calificacion, 15)
            sCadImp = sCadImp & ImpreFormat(Format(R!Vigencia, "dd/mm/yyyy"), 10)
            sCadImp = sCadImp & ImpreFormat(Format(R!Vencimiento, "dd/mm/yyyy"), 10)
            sCadImp = sCadImp & ImpreFormat(Format(R!Pagado, "dd/mm/yyyy"), 10)
            sCadImp = sCadImp & ImpreFormat(Format(R!Proxima, "dd/mm/yyyy"), 10)
            sCadImp = sCadImp & ImpreFormat(R!Mora, 10, 0)
            sCadImp = sCadImp & ImpreFormat(R!Tasa, 5, 2)
            sCadImp = sCadImp & ImpreFormat(R!Int_Devengado, 12, 2)
            sCadImp = sCadImp & ImpreFormat(R!Int_Suspenso, 12, 2)
            sCadImp = sCadImp & ImpreFormat(R!Provision, 12, 2)
            sCadImp = sCadImp & ImpreFormat(R!Tasa_Provision, 12, 2)
            sCadImp = sCadImp & oImpresora.gPrnSaltoLinea
            
            lnTotAAprobado = lnTotAAprobado + R!Monto_Aprobado
            lnTotASaldoK = lnTotASaldoK + R!Saldo_Capital
            lnTotADevengado = lnTotADevengado + R!Int_Devengado
            lnTotASuspenso = lnTotASuspenso + R!Int_Suspenso
            lnTotAProvision = lnTotAProvision + R!Provision
            
            lnTotGAprobado = lnTotGAprobado + R!Monto_Aprobado
            lnTotGSaldoK = lnTotGSaldoK + R!Saldo_Capital
            lnTotGDevengado = lnTotGDevengado + R!Int_Devengado
            lnTotGSuspenso = lnTotGSuspenso + R!Int_Suspenso
            lnTotGProvision = lnTotGProvision + R!Provision
            
            
        
            If j Mod 500 = 0 Then
                lsCadImp = lsCadImp & sCadImp
                sCadImp = ""
            End If
            If K = 45 Then
                sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
                sw = 0
                K = 0
            End If
            R.MoveNext
            
            
            If R.EOF Then
                Exit Do
          
            End If
       Loop
         
         Call ImprimeTotalesCarteraDetxIntDSP(sCadImp, lnTotAAprobado, lnTotASaldoK, lnTotADevengado, lnTotASuspenso, lnTotAProvision, lnCuantos, "Sub Total Por Grupo: ")
         sCadImp = sCadImp & oImpresora.gPrnSaltoLinea
         sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
         
'         Call ImprimeTotalesCarteraDetxIntDSP(sCadImp, lnTotAAprobado, lnTotASaldoK, lnTotADevengado, lnTotASuspenso, lnTotAProvision, lnCuantos, "Sub Total Segun Tipo: ")
'         sCadImp = sCadImp & oImpresora.gPrnSaltoLinea
'         Call ImprimeTotalesCarteraDetxIntDSP(sCadImp, lnTotAAprobado, lnTotASaldoK, lnTotADevengado, lnTotASuspenso, lnTotAProvision, lnCuantos, "Sub Total Segun Moneda: ")
'         sCadImp = sCadImp & oImpresora.gPrnSaltoLinea
'         Call ImprimeTotalesCarteraDetxIntDSP(sCadImp, lnTotAAprobado, lnTotASaldoK, lnTotADevengado, lnTotASuspenso, lnTotAProvision, lnCuantos, "Sub Total Segun Estado: ")
'
         
         
         lnCuantos = 0
         sw = 2
         
    Loop
    Call ImprimeTotalesCarteraDetxIntDSP(sCadImp, lnTotGAprobado, lnTotGSaldoK, lnTotGDevengado, lnTotGSuspenso, lnTotGProvision, lnGCuantos, "Total General: ")
    R.Close
    Set R = Nothing
    sCadImp = sCadImp & oImpresora.gPrnCondensadaOFF
    ImprimeCarteraDetxIntDSP = lsCadImp & sCadImp
End Function

'By MAVM 27112008
Public Sub GeneraReporteEnArchivoExcel(ByVal psNomCmac As String, ByVal psNomAge As String, ByVal psCodUser As String, ByVal pdFecsis As Date, ByVal psTitulo As String, ByVal psSubTitulo As String, _
                                    ByVal psNomArchivo As String, ByVal pMatCabeceras As Variant, ByVal prRegistros As ADODB.Recordset, _
                                    Optional pnNumDecimales As Integer, Optional Visible As Boolean = False, Optional psNomHoja As String = "", _
                                    Optional pbSinFormatDeReg As Boolean = False, _
                                    Optional pbUsarCabecerasDeRS As Boolean = False)
    Dim rs As ADODB.Recordset
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim liLineas As Integer, I As Integer
    Dim fs As Scripting.FileSystemObject
    Dim lnNumColumns As Integer

    Dim lnTotalAprobado As Double
    Dim lnTotalSaldok As Double
    Dim lnTotalDevengado As Double
    Dim lnTotalSuspenso As Double
    Dim lnTotalProvision As Double
    
    lnTotalAprobado = 0
    lnTotalSaldok = 0
    lnTotalDevengado = 0
    lnTotalSuspenso = 0
    lnTotalProvision = 0
    
     Do While Not prRegistros.EOF
        lnTotalAprobado = lnTotalAprobado + prRegistros!Monto_Aprobado
        lnTotalSaldok = lnTotalSaldok + prRegistros!Saldo_Capital
        lnTotalDevengado = lnTotalDevengado + prRegistros!Int_Devengado
        lnTotalSuspenso = lnTotalSuspenso + prRegistros!Int_Suspenso
        lnTotalProvision = lnTotalProvision + prRegistros!Provision
        prRegistros.MoveNext
    Loop

    If Not (prRegistros.EOF And prRegistros.BOF) Then
        If pbUsarCabecerasDeRS = True Then
            lnNumColumns = prRegistros.Fields.Count
        Else
            lnNumColumns = UBound(pMatCabeceras)
            lnNumColumns = IIf(prRegistros.Fields.Count < lnNumColumns, prRegistros.Fields.Count, prRegistros.Fields.Count)
        End If

        If psNomHoja = "" Then psNomHoja = psNomArchivo
        psNomArchivo = psNomArchivo & "_" & psCodUser & ".xls"

        Set fs = New Scripting.FileSystemObject
        Set xlAplicacion = New Excel.Application
        If fs.FileExists(App.Path & "\Spooler\" & psNomArchivo) Then
            fs.DeleteFile (App.Path & "\Spooler\" & psNomArchivo)
        End If
        Set xlLibro = xlAplicacion.Workbooks.Add
        Set xlHoja1 = xlLibro.Worksheets.Add

        xlHoja1.Name = psNomHoja
        xlHoja1.Cells.Select
        
        'Cabeceras
        xlHoja1.Cells(1, 1) = psNomCmac
        xlHoja1.Cells(1, lnNumColumns) = Trim(Format(pdFecsis, "dd/mm/yyyy hh:mm:ss"))
        xlHoja1.Cells(2, 1) = psNomAge
        xlHoja1.Cells(2, lnNumColumns) = psCodUser
        xlHoja1.Cells(4, 1) = psTitulo
        xlHoja1.Cells(5, 1) = psSubTitulo
        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, lnNumColumns)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(5, lnNumColumns)).HorizontalAlignment = xlCenter

        liLineas = 6
        If pbUsarCabecerasDeRS = True Then
            For I = 0 To prRegistros.Fields.Count - 1
                xlHoja1.Cells(liLineas, I + 1) = prRegistros.Fields(I).Name
            Next I
        Else
            For I = 0 To lnNumColumns - 1
                If (I + 1) > UBound(pMatCabeceras) Then
                    xlHoja1.Cells(liLineas, I + 1) = prRegistros.Fields(I).Name
                Else
                    xlHoja1.Cells(liLineas, I + 1) = pMatCabeceras(I, 0)
                End If
            Next I
        End If
        
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).Cells.Interior.Color = RGB(220, 220, 220)
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).HorizontalAlignment = xlCenter

        If pbSinFormatDeReg = False Then
            liLineas = liLineas + 1
            While Not prRegistros.EOF
                For I = 0 To lnNumColumns - 1
                    If pMatCabeceras(I, 1) = "" Then  'Verificamos si tiene tipo
                        xlHoja1.Cells(liLineas, I + 1) = prRegistros(I)
                    Else
                        Select Case pMatCabeceras(I, 1)
                            Case "S"
                                xlHoja1.Cells(liLineas, I + 1) = prRegistros(I)
                            Case "N"
                                xlHoja1.Cells(liLineas, I + 1) = Format(prRegistros(I), "#0.00")
                            Case "D"
                                xlHoja1.Cells(liLineas, I + 1) = IIf(Format(prRegistros(I), "yyyymmdd") = "19000101", "", Format(prRegistros(I), "dd/mm/yyyy"))
                        End Select
                    End If
                Next I
                liLineas = liLineas + 1
                prRegistros.MoveNext
            Wend
        Else
            xlHoja1.Range("A7").CopyFromRecordset prRegistros 'Copia el contenido del recordset a excel
        End If
        
        xlHoja1.Range(xlHoja1.Cells(prRegistros.RecordCount + 8, 15), xlHoja1.Cells(prRegistros.RecordCount + 8, 21)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(prRegistros.RecordCount + 8, 15), xlHoja1.Cells(prRegistros.RecordCount + 8, 21)).ColumnWidth = 16
        xlHoja1.Cells(prRegistros.RecordCount + 8, 15) = "TOTALES"
        xlHoja1.Cells(prRegistros.RecordCount + 8, 16) = Format(lnTotalAprobado, "#,###,###.##")
        xlHoja1.Cells(prRegistros.RecordCount + 8, 17) = Format(lnTotalSaldok, "#,###,###.##")
        xlHoja1.Cells(prRegistros.RecordCount + 8, 18) = Format(lnTotalDevengado, "#,###,###.##")
        xlHoja1.Cells(prRegistros.RecordCount + 8, 19) = Format(lnTotalSuspenso, "#,###,###.##")
        xlHoja1.Cells(prRegistros.RecordCount + 8, 20) = Format(lnTotalProvision, "#,###,###.##")
       
        xlHoja1.SaveAs App.Path & "\Spooler\" & psNomArchivo
        MsgBox "Se ha generado el Archivo en " & App.Path & "\Spooler\" & psNomArchivo

        If Visible Then
            xlAplicacion.Visible = True
            xlAplicacion.Windows(1).Visible = True
        End If

        'xlLibro.Close
        'xlAplicacion.Quit
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing

    End If
End Sub

'By Capi 08022008
Public Function ImprimeCarteraResxLineas(pnTipCambio As Double, ByVal psNomCmac As String, _
                                         ByVal psNomAge As String, ByVal pdFecsis As Date, _
                                         ByVal psCodUser As String, _
                                         Optional psExcel As Boolean = False)

  
    Dim lsFecha As String
    Dim lnPeriodo As Integer
    'By capi 17112008
    Dim lnMes As Integer
    Dim lnAnno As Integer
    '
    Dim lsMoneda As String
    Dim R As New ADODB.Recordset
    Dim oDCre As DCreditos
    Dim vAgencia As String
    Dim vEstado As String
    Dim vMoneda As String
    Dim vCadena As String
    
    Dim vTipo As String
    Dim vSituacion As String
    Dim sCadImp As String
    
    
    Dim j As Double
    Dim K As Integer
    Dim lnCuantos As Integer
    Dim lnGCuantos As Integer
    Dim sw As Integer
    
    Dim lnTotAAprobado As Double
    Dim lnTotASaldoK As Double
    Dim lnTotADevengado As Double
    Dim lnTotASuspenso As Double
    Dim lnTotAProvision As Double
    
    Dim lnTotGAprobado As Double
    Dim lnTotGSaldoK As Double
    Dim lnTotGDevengado As Double
    Dim lnTotGSuspenso As Double
    Dim lnTotGProvision As Double
    
    
    
        
    
    
    'By capi 17112008
    lnMes = Month(pdFecsis)
    lnAnno = Year(pdFecsis)
    '
    
    
    Set oDCre = New DCreditos
    'By capi 17112008
    'Set R = oDCre.ObtenerCalCarteraxLineas("1", 0, 0)
    Set R = oDCre.ObtenerCalCarteraxLineas("1", 0, 0, lnMes, lnAnno)
    
    Set oDCre = Nothing
    lsMoneda = "Expresado en Moneda Nacional al T.C. : "
        
        
    lsFecha = "Periodo: " & lnMes & "/" & lnAnno
    
    If R!Cuantos = 0 Then
        MsgBox "Cierre de Calificacion Pendiente - No Existe Registros a Procesar", vbInformation, "Aviso"
        R.Close: Set R = Nothing
        Exit Function
    Else
        MsgBox "Se Verifico Cierre Calificacion - Registros a Procesar >>>>>> " & R!Cuantos, vbInformation, "Aviso"
    End If

        
        
    
    Set oDCre = New DCreditos
    'By capi 17112008
    'Set R = oDCre.ObtenerCalCarteraxLineas("2", lnPeriodo, pnTipCambio)
    Set R = oDCre.ObtenerCalCarteraxLineas("2", 0, pnTipCambio, lnMes, lnAnno)
    '
    Set oDCre = Nothing
    
   If psExcel = True Then
        Dim lMatCabecera As Variant
        ReDim lMatCabecera(13, 2)
        lMatCabecera(0, 0) = "Estado": lMatCabecera(0, 1) = ""
        lMatCabecera(1, 0) = "Moneda": lMatCabecera(1, 1) = ""
        lMatCabecera(2, 0) = "Condicion": lMatCabecera(2, 1) = ""
        lMatCabecera(3, 0) = "Tipo": lMatCabecera(3, 1) = ""
        lMatCabecera(4, 0) = "cLinea": lMatCabecera(4, 1) = ""
        lMatCabecera(5, 0) = "Agencia": lMatCabecera(5, 1) = ""
        lMatCabecera(6, 0) = "Linea": lMatCabecera(6, 1) = ""
        lMatCabecera(7, 0) = "Descripcion": lMatCabecera(7, 1) = ""
        lMatCabecera(8, 0) = "Monto_Aprobado": lMatCabecera(8, 1) = ""
        lMatCabecera(9, 0) = "Saldo_Capital": lMatCabecera(9, 1) = ""
        lMatCabecera(10, 0) = "Int_Devengado": lMatCabecera(10, 1) = ""
        lMatCabecera(11, 0) = "Int_Suspenso": lMatCabecera(11, 1) = ""
        lMatCabecera(12, 0) = "Provision": lMatCabecera(12, 1) = ""
                
        Call GeneraReporteEnArchivoExcelResumen(gsNomCmac, gsNomAge, psCodUser, gdFecSis, "RESUMEN DE CARTERA POR LINEAS (CALIFICACION) - " & lsFecha, "", "ResumenCarteraXLinea", lMatCabecera, R, 2, True, , True)
        Exit Function
    End If
    
       
    sCadImp = ""
    
    
        

    
    j = 0
    sw = 0
    K = 0
    lnCuantos = 0
    lnGCuantos = 0
    
    lnTotGAprobado = 0
    lnTotGSaldoK = 0
    lnTotGDevengado = 0
    lnTotGSuspenso = 0
    lnTotGProvision = 0
    
              
    
    
    Do While Not R.EOF
       
        vEstado = R!Estado
        vMoneda = R!Moneda
        vSituacion = R!Condicion
       
        vCadena = R!Estado & R!Moneda & R!Condicion
        
        lnTotAAprobado = 0
        lnTotASaldoK = 0
        lnTotADevengado = 0
        lnTotASuspenso = 0
        lnTotAProvision = 0
        
        Do While R!Estado & R!Moneda & R!Condicion = vCadena
            j = j + 1
            K = K + 1
            lnCuantos = lnCuantos + 1
            lnGCuantos = lnGCuantos + 1
            If sw = 0 Or sw = 2 Then
                sw = 1
                'sCadImp = sCadImp & oImpresora.gPrnCondensadaON
                Call ImprimeCabeceraDocumentoAncho(sCadImp, psNomAge, pdFecsis, psCodUser, "RESUMEN DE CARTERA POR LINEAS (CALIFICACION) - " & lsFecha, psNomCmac, , True, , lsMoneda & pnTipCambio)
                
                sCadImp = sCadImp & String(200, "-") & oImpresora.gPrnSaltoLinea
                '                    0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
                '                              1         2         3         4         5         6         7         8         9         10        11        12        13        14        15        16        17
                sCadImp = sCadImp & " AGEN.     TIPO                    CONDICION           LINEA        DESCRIPCION                                            MONTO APRO.     SALDO K    DEVENGADO       SUSPENSO        PROVISION  " & oImpresora.gPrnSaltoLinea
                sCadImp = sCadImp & String(200, "-") & oImpresora.gPrnSaltoLinea
                
                
                sCadImp = sCadImp & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                sCadImp = sCadImp & R!Estado & oImpresora.gPrnSaltoLinea
                If R!Moneda = "1" Then
                    sCadImp = sCadImp & "Moneda Nacional" & oImpresora.gPrnSaltoLinea
                Else
                    sCadImp = sCadImp & "Moneda Extranjera" & oImpresora.gPrnSaltoLinea
                End If
                
                sCadImp = sCadImp & R!Condicion & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            End If
            
                
            'sCadImp = sCadImp & ImpreFormat(J, 4, 0)
            
            sCadImp = sCadImp & ImpreFormat(R!Agencia, 3)
            sCadImp = sCadImp & ImpreFormat(R!Tipo, 30)
            sCadImp = sCadImp & ImpreFormat(R!Condicion, 15)
            sCadImp = sCadImp & ImpreFormat(R!Linea, 10)
            sCadImp = sCadImp & ImpreFormat(R!Descripcion, 50)
            sCadImp = sCadImp & ImpreFormat(R!Monto_Aprobado, 12, 2)
            sCadImp = sCadImp & ImpreFormat(R!Saldo_Capital, 12, 2)
            sCadImp = sCadImp & ImpreFormat(R!Int_Devengado, 12, 2)
            sCadImp = sCadImp & ImpreFormat(R!Int_Suspenso, 12, 2)
            sCadImp = sCadImp & ImpreFormat(R!Provision, 12, 2)
            sCadImp = sCadImp & oImpresora.gPrnSaltoLinea
            
            lnTotAAprobado = lnTotAAprobado + R!Monto_Aprobado
            lnTotASaldoK = lnTotASaldoK + R!Saldo_Capital
            lnTotADevengado = lnTotADevengado + R!Int_Devengado
            lnTotASuspenso = lnTotASuspenso + R!Int_Suspenso
            lnTotAProvision = lnTotAProvision + R!Provision
            
            lnTotGAprobado = lnTotGAprobado + R!Monto_Aprobado
            lnTotGSaldoK = lnTotGSaldoK + R!Saldo_Capital
            lnTotGDevengado = lnTotGDevengado + R!Int_Devengado
            lnTotGSuspenso = lnTotGSuspenso + R!Int_Suspenso
            lnTotGProvision = lnTotGProvision + R!Provision
            
            
            If K = 45 Then
                
                sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
                sw = 0
                K = 0
            End If
            R.MoveNext
            
            
            If R.EOF Then
                Exit Do
          
            End If
       Loop
         
         Call ImprimeTotalesCarteraDetxIntDSP(sCadImp, lnTotAAprobado, lnTotASaldoK, lnTotADevengado, lnTotASuspenso, lnTotAProvision, lnCuantos, "Sub Total Por Grupo: ")
         sCadImp = sCadImp & oImpresora.gPrnSaltoLinea
         
         sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
         
         lnCuantos = 0
         sw = 2
         K = 0
    Loop
    Call ImprimeTotalesCarteraDetxIntDSP(sCadImp, lnTotGAprobado, lnTotGSaldoK, lnTotGDevengado, lnTotGSuspenso, lnTotGProvision, lnGCuantos, "Total General: ")

    R.Close
    Set R = Nothing
    'sCadImp = sCadImp & oImpresora.gPrnCondensadaOFF
    ImprimeCarteraResxLineas = sCadImp
End Function

'By MAVM 27112008
Public Sub GeneraReporteEnArchivoExcelResumen(ByVal psNomCmac As String, ByVal psNomAge As String, ByVal psCodUser As String, ByVal pdFecsis As Date, ByVal psTitulo As String, ByVal psSubTitulo As String, _
                                    ByVal psNomArchivo As String, ByVal pMatCabeceras As Variant, ByVal prRegistros As ADODB.Recordset, _
                                    Optional pnNumDecimales As Integer, Optional Visible As Boolean = False, Optional psNomHoja As String = "", _
                                    Optional pbSinFormatDeReg As Boolean = False, _
                                    Optional pbUsarCabecerasDeRS As Boolean = False)
    Dim rs As ADODB.Recordset
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim liLineas As Integer, I As Integer
    Dim fs As Scripting.FileSystemObject
    Dim lnNumColumns As Integer

    Dim lnTotalAprobado As Double
    Dim lnTotalSaldok As Double
    Dim lnTotalDevengado As Double
    Dim lnTotalSuspenso As Double
    Dim lnTotalProvision As Double
    
    lnTotalAprobado = 0
    lnTotalSaldok = 0
    lnTotalDevengado = 0
    lnTotalSuspenso = 0
    lnTotalProvision = 0
    
     Do While Not prRegistros.EOF
        lnTotalAprobado = lnTotalAprobado + prRegistros!Monto_Aprobado
        lnTotalSaldok = lnTotalSaldok + prRegistros!Saldo_Capital
        lnTotalDevengado = lnTotalDevengado + prRegistros!Int_Devengado
        lnTotalSuspenso = lnTotalSuspenso + prRegistros!Int_Suspenso
        lnTotalProvision = lnTotalProvision + prRegistros!Provision
        prRegistros.MoveNext
    Loop
    
    If Not (prRegistros.EOF And prRegistros.BOF) Then
        If pbUsarCabecerasDeRS = True Then
            lnNumColumns = prRegistros.Fields.Count
        Else
            lnNumColumns = UBound(pMatCabeceras)
            lnNumColumns = IIf(prRegistros.Fields.Count < lnNumColumns, prRegistros.Fields.Count, prRegistros.Fields.Count)
        End If

        If psNomHoja = "" Then psNomHoja = psNomArchivo
        psNomArchivo = psNomArchivo & "_" & psCodUser & ".xls"

        Set fs = New Scripting.FileSystemObject
        Set xlAplicacion = New Excel.Application
        If fs.FileExists(App.Path & "\Spooler\" & psNomArchivo) Then
            fs.DeleteFile (App.Path & "\Spooler\" & psNomArchivo)
        End If
        Set xlLibro = xlAplicacion.Workbooks.Add
        Set xlHoja1 = xlLibro.Worksheets.Add

        xlHoja1.Name = psNomHoja
        xlHoja1.Cells.Select
        
        'Cabeceras
        xlHoja1.Cells(1, 1) = psNomCmac
        xlHoja1.Cells(1, lnNumColumns) = Trim(Format(pdFecsis, "dd/mm/yyyy hh:mm:ss"))
        xlHoja1.Cells(2, 1) = psNomAge
        xlHoja1.Cells(2, lnNumColumns) = psCodUser
        xlHoja1.Cells(4, 1) = psTitulo
        xlHoja1.Cells(5, 1) = psSubTitulo
        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, lnNumColumns)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(5, lnNumColumns)).HorizontalAlignment = xlCenter

        liLineas = 6
        If pbUsarCabecerasDeRS = True Then
            For I = 0 To prRegistros.Fields.Count - 1
                xlHoja1.Cells(liLineas, I + 1) = prRegistros.Fields(I).Name
            Next I
        Else
            For I = 0 To lnNumColumns - 1
                If (I + 1) > UBound(pMatCabeceras) Then
                    xlHoja1.Cells(liLineas, I + 1) = prRegistros.Fields(I).Name
                Else
                    xlHoja1.Cells(liLineas, I + 1) = pMatCabeceras(I, 0)
                End If
            Next I
        End If
        
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).Cells.Interior.Color = RGB(220, 220, 220)
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).HorizontalAlignment = xlCenter

        If pbSinFormatDeReg = False Then
            liLineas = liLineas + 1
            While Not prRegistros.EOF
                For I = 0 To lnNumColumns - 1
                    If pMatCabeceras(I, 1) = "" Then  'Verificamos si tiene tipo
                        xlHoja1.Cells(liLineas, I + 1) = prRegistros(I)
                    Else
                        Select Case pMatCabeceras(I, 1)
                            Case "S"
                                xlHoja1.Cells(liLineas, I + 1) = prRegistros(I)
                            Case "N"
                                xlHoja1.Cells(liLineas, I + 1) = Format(prRegistros(I), "#0.00")
                            Case "D"
                                xlHoja1.Cells(liLineas, I + 1) = IIf(Format(prRegistros(I), "yyyymmdd") = "19000101", "", Format(prRegistros(I), "dd/mm/yyyy"))
                        End Select
                    End If
                Next I
                liLineas = liLineas + 1
                prRegistros.MoveNext
            Wend
        Else
            xlHoja1.Range("A7").CopyFromRecordset prRegistros 'Copia el contenido del recordset a excel
        End If
        
        xlHoja1.Range(xlHoja1.Cells(prRegistros.RecordCount + 8, 8), xlHoja1.Cells(prRegistros.RecordCount + 8, 13)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(prRegistros.RecordCount + 8, 8), xlHoja1.Cells(prRegistros.RecordCount + 8, 13)).ColumnWidth = 16
        xlHoja1.Cells(prRegistros.RecordCount + 8, 8) = "TOTALES"
        xlHoja1.Cells(prRegistros.RecordCount + 8, 9) = Format(lnTotalAprobado, "#,###,###.##")
        xlHoja1.Cells(prRegistros.RecordCount + 8, 10) = Format(lnTotalSaldok, "#,###,###.##")
        xlHoja1.Cells(prRegistros.RecordCount + 8, 11) = Format(lnTotalDevengado, "#,###,###.##")
        xlHoja1.Cells(prRegistros.RecordCount + 8, 12) = Format(lnTotalSuspenso, "#,###,###.##")
        xlHoja1.Cells(prRegistros.RecordCount + 8, 13) = Format(lnTotalProvision, "#,###,###.##")
       
        xlHoja1.SaveAs App.Path & "\Spooler\" & psNomArchivo
        MsgBox "Se ha generado el Archivo en " & App.Path & "\Spooler\" & psNomArchivo

        If Visible Then
            xlAplicacion.Visible = True
            xlAplicacion.Windows(1).Visible = True
        End If

            'xlLibro.Close
            'xlAplicacion.Quit
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing

    End If
End Sub
'EJVG 20110902*******************************
Public Function validarGeneracionRepTripleJump(pdFechaFinMes As Date) As Boolean
    Dim oConect As DConecta
    Dim sSQL As String
    Dim rs As Recordset
    Set oConect = New DConecta
    Set rs = New Recordset
    sSQL = "exec stp_sel_ValidaGeneracionRepTripleJump '" & Format(pdFechaFinMes, "yyyymmdd") & "'"
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(sSQL)
    validarGeneracionRepTripleJump = rs!Exito
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function obtenerPlazoFijoxRangos(ByVal pdFecha As Date, ByVal pnFiltro As Integer) As Recordset
    Dim oConect As DConecta
    Dim sSQL As String
    Set oConect = New DConecta
    sSQL = "exec stp_sel_PlazoFijoRangos '" & Format(pdFecha, "yyyymmdd") & "'," & pnFiltro
    oConect.AbreConexion
    Set obtenerPlazoFijoxRangos = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function obtenerCarteraxRangoTripleJump(ByVal pdFecha As Date) As Recordset
    Dim oConect As DConecta
    Dim sSQL As String
    Set oConect = New DConecta
    sSQL = "exec stp_sel_CarteraxRangoTripleJump '" & Format(pdFecha, "yyyymmdd") & "'"
    oConect.AbreConexion
    Set obtenerCarteraxRangoTripleJump = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function obtenerCarteraCastigadaxRangoAño(ByVal pdFechaFinMes As Date) As Recordset
    Dim oConect As DConecta
    Dim sSQL As String
    Set oConect = New DConecta
    sSQL = "exec stp_sel_ObtenerCarteraCastigadaxRangoAño '" & Format(pdFechaFinMes, "yyyymmdd") & "'"
    oConect.AbreConexion
    Set obtenerCarteraCastigadaxRangoAño = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function obtenerEstadisticaOperacionTripleJump(ByVal pdFechaFinMes As Date) As Recordset
    Dim oConect As DConecta
    Dim sSQL As String
    Set oConect = New DConecta
    sSQL = "exec stp_sel_obtenerEstadisticaOperacionesTripleJump '" & Format(pdFechaFinMes, "yyyymmdd") & "'"
    oConect.AbreConexion
    Set obtenerEstadisticaOperacionTripleJump = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function obtenerTotalMargenFormaAYB(ByVal pnTpoInfo As Integer, ByVal pdFechaFinMes As Date, psCodigo As String) As Currency
    Dim lnCampo As String
    Select Case pnTpoInfo
        Case 1 'Moneda Nacional
            lnCampo = "nMN"
        Case 2 'Moneda Extranjera
            lnCampo = "nME"
        Case 3 'Total de 1 y 2
            lnCampo = "nTotal"
        Case Else
            obtenerTotalMargenFormaAYB = 0
            Exit Function
    End Select
    
    Dim oConect As DConecta
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    Set oConect = New DConecta
    
    sSQL = "select " & lnCampo & " from BalanceGen WHERE cFecha = '" & Format(pdFechaFinMes, "yyyymmdd") & "' And cBalanceCate = '3' and cCodigo='" & psCodigo & "'"
    oConect.AbreConexion
     
    Set rs = oConect.CargaRecordSet(sSQL)
    If rs.EOF And rs.BOF Then
        obtenerTotalMargenFormaAYB = 0
    Else
        If IsNull(rs.Fields(0)) Then
            obtenerTotalMargenFormaAYB = 0
        Else
            obtenerTotalMargenFormaAYB = rs.Fields(0)
        End If
    End If
    oConect.CierraConexion
    Set rs = Nothing
    Set oConect = Nothing
End Function
Public Function obtenerTotalImporteActivoNuevo(psFechaFinMes As Date) As Currency
    Dim oConect As DConecta
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim lnImporte As Currency
    Dim I As Integer
    
    Set oConect = New DConecta
    Set rs = New ADODB.Recordset
    
    'sSql = "exec stp_sel_RS760107_DetalleComprasActivosFijos '" & Format(psFechaFinMes, "dd/mm/yyyy") & "','1'"
    'EJVG20120515
    sSQL = "exec stp_sel_RS760107_DetalleComprasActivosFijos '" & Format(CDate("01/" & Format(Month(psFechaFinMes), "00") & "/" & Year(psFechaFinMes)), "dd/mm/yyyy") & "','" & Format(psFechaFinMes, "dd/mm/yyyy") & "'"
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(sSQL)
    
    lnImporte = 0
    If Not rs.EOF Then
        For I = 0 To rs.RecordCount - 1
            lnImporte = lnImporte + CCur(rs!nMovImporte)
            rs.MoveNext
        Next
    End If
    obtenerTotalImporteActivoNuevo = lnImporte
    oConect.CierraConexion
    Set rs = Nothing
    Set oConect = Nothing
End Function

Public Sub obtenerCortoLargoDepositoPlazoFijo(ByVal psFechaFinMes As Date, ByVal pnTipoCambio As Currency, ByRef pnCP_DPF As Currency, ByRef pnLP_DPF As Currency)
    Dim oConect As DConecta
    Dim sSQL As String
    Dim rs As ADODB.Recordset

    Set oConect = New DConecta
    Set rs = New ADODB.Recordset
    
    sSQL = "exec stp_sel_ReporteAhorro '" & Format(psFechaFinMes, "yyyymmdd") & "'," & pnTipoCambio
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(sSQL)

    Do While Not rs.EOF
        If rs!nPlazo = 1 Then
            pnCP_DPF = pnCP_DPF + (CCur(rs!nSaldoPF) + CCur(rs!nSaldoPFRestr) + CCur(rs!nSaldoPlazFiEmSisFi)) * IIf(rs!cMoneda = 1, 1, pnTipoCambio)
        ElseIf rs!nPlazo = 2 Then
            pnLP_DPF = pnLP_DPF + (CCur(rs!nSaldoPF) + CCur(rs!nSaldoPFRestr)) * IIf(rs!cMoneda = 1, 1, pnTipoCambio)
        End If
        rs.MoveNext
    Loop
    oConect.CierraConexion
    Set rs = Nothing
    Set oConect = Nothing
End Sub
'********************************************
Public Function PlazoFijoxRango(ByVal psFecha As Date, ByVal pnMoneda As Integer) As String
    On Error GoTo InteresCTSxFxAGErr
    Dim sSQL As String
    Dim sCabecera As String
    Dim sAsiento As String
    Dim nLi As Long
    Dim lsFechaIni As String
    Dim lnMoneda As String
    Dim lnSaldo As Currency
        
    Dim rs       As New ADODB.Recordset
    Dim oConect As DConecta
    Set oConect = New DConecta
    
    If oConect.AbreConexion = False Then Exit Function
    
    sCabecera = ""
    sAsiento = ""
    nLi = 0
    
    lsFechaIni = Format(psFecha, "yyyymmdd")
    lnMoneda = pnMoneda

     sSQL = " select 'HASTA 30 DIAS' cRango, ISNULL(SUM(nSaldCnt),0) nSaldo" & _
            " From capsaldosdiarios" & _
            " where datediff(day,dfecha,'" & lsFechaIni & "' )=0 and nplazo <=30" & _
            " And substring(cctacod,6,3)='233' and substring(cctacod,9,1)=" & lnMoneda & _
            " Union All " & _
            " select 'DE 31 A 90 DIAS' cRango, SUM(nSaldCnt) nSaldo" & _
            " From capsaldosdiarios" & _
            " where datediff(day,dfecha,'" & lsFechaIni & "')=0 and nplazo >30 and nplazo <=90" & _
            " And substring(cctacod,6,3)='233' and substring(cctacod,9,1)=" & lnMoneda
            
    sSQL = sSQL & " Union ALL" & _
           " select 'DE 91 A 180 DIAS' cRango, SUM(nSaldCnt) nSaldo" & _
           " From capsaldosdiarios" & _
           " where datediff(day,dfecha,'" & lsFechaIni & "')=0 and nplazo >90 and nplazo <=180" & _
           " And substring(cctacod,6,3)='233' and substring(cctacod,9,1)=" & lnMoneda & _
           " Union All" & _
           " select 'DE 181 A 360 DIAS' cRango, SUM(nSaldCnt) nSaldo" & _
           " From capsaldosdiarios" & _
           " where datediff(day,dfecha,'" & lsFechaIni & "')=0 and nplazo >180 and nplazo <=360" & _
           " And substring(cctacod,6,3)='233' and substring(cctacod,9,1)=" & lnMoneda

    sSQL = sSQL & " UNION ALL" & _
           " select 'DE 361 A 540 DIAS' cRango, SUM(nSaldCnt) nSaldo" & _
           " From capsaldosdiarios" & _
           " where datediff(day,dfecha,'" & lsFechaIni & "')=0 and nplazo >360 and nplazo <=540" & _
           " And substring(cctacod,6,3)='233' and substring(cctacod,9,1)=" & lnMoneda & _
           " Union All" & _
           " select 'DE 541 A MAS DIAS' cRango, SUM(nSaldCnt) nSaldo" & _
           " From capsaldosdiarios" & _
           " where datediff(day,dfecha,'" & lsFechaIni & "')=0 and nplazo >540" & _
           " And substring(cctacod,6,3)='233' and substring(cctacod,9,1)=" & lnMoneda

    
       Set rs = oConect.CargaRecordSet(sSQL)
       If rs.EOF Then
          Exit Function
       End If
         
       Dim MonedaDesc As String
       If lnMoneda = 1 Then
          '''MonedaDesc = "SOLES" 'MARG ERS044-2016
          MonedaDesc = StrConv(gcPEN_PLURAL, vbUpperCase) 'MARG ERS044-2016
       Else
          MonedaDesc = "DOLARES"
       End If
    
    sCabecera = sCabecera & "FECHA:" & Space(3) & lsFechaIni & Space(10) & "MONEDA:" & Space(3) & MonedaDesc & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera & "RANGO" & Space(25) & "MONTO" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera & String(43, "_") & oImpresora.gPrnSaltoLinea
    
    Dim Rango As String * 20
    Dim Monto As String * 20
    lnSaldo = 0
    Do While Not rs.EOF
       DoEvents
       If nLi > 60 Then
           nLi = 5
           sAsiento = sAsiento & oImpresora.gPrnSaltoPagina & sCabecera
       End If
       If rs!cRango = "HASTA 30 DIAS" Or rs!cRango = "DE 31 A 90 DIAS" Or rs!cRango = "DE 91 A 180 DIAS" Or rs!cRango = "DE 181 A 360 DIAS" Then
            lnSaldo = lnSaldo + IIf(IsNull(rs!nSaldo), 0, rs!nSaldo)
       End If
       Rango = rs!cRango
       Monto = IIf(IsNull(rs!nSaldo), 0, rs!nSaldo)
       sAsiento = sAsiento & Rango & Space(10) & Monto & oImpresora.gPrnSaltoLinea
          
       nLi = nLi + 1
       rs.MoveNext
    Loop
    Dim oDBalan As DbalanceCont
    Set oDBalan = New DbalanceCont
    Call oDBalan.InsertaCtaContSaldoDiario("21" & CStr(pnMoneda) & "3", psFecha, "763405", lnSaldo)
    Set oDBalan = Nothing
    sAsiento = sCabecera & sAsiento
    PlazoFijoxRango = sAsiento
    Exit Function
InteresCTSxFxAGErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:InteresCTSxFxAG Method")
End Function

Public Function ImprimeAsientoContableProv(ByVal psMovNro As String, ByVal pnLinPage As Integer, ByVal pnColPage As Integer, Optional ByVal psTit1 As String, Optional ByVal psCabe1 As String, Optional ByVal psPiePag As String = "9", Optional ByVal psOtraFirma As String = "", Optional psNomCmact As String = "") As String
    On Error GoTo ImprimeAsientoContableErr
    Dim sSQL As String
    Dim rs       As New ADODB.Recordset
    Dim rsDoc       As New ADODB.Recordset
    Dim nTot     As Currency
    Dim nTotH    As Currency
    Dim sDoc     As String
    Dim sTexto   As String
    Dim sAsiento As String
    Dim nLi      As Integer
    Dim sCabecera As String
    Dim lnCambioFijo As Currency
    Dim lnCambioVenta As Currency
    Dim lnCambioCompra As Currency
    Dim lnCambioVentaSBS As Currency
    Dim lnCambioCompraSBS As Currency
    Dim CON As String
    Dim COFF As String
    Dim lsMovBilletaje As String
    
    Dim lsDesObjeto As String
    Dim lsFecha    As String
    Dim oConect As DConecta
    Dim lsGlosa As String
    Dim lsMovNro As String
    Dim lsOpeCod As String
    Dim sDestino As String
    Dim sOrigen As String
    Dim sDNI As String
    Dim sRUC As String
    Dim sCarnet As String
    Dim oDGeneral As DGeneral
    Dim BON As String
    Dim BOFF As String
    Dim oTipoCamb As nTipoCambio
    
    'TORE 20190801 - RFC1712260009
    Dim rsTC As ADODB.Recordset
    Dim sqlTC As String
    Dim lsTCFechaDoc  As String
    'END TORE 20190801 - RFC1712260009
    
    Dim sAsiento8 As String, nTot8 As Double, nTotH8 As Double, sAsiento8c As String
    
    Set oDGeneral = New DGeneral
    Set oTipoCamb = New nTipoCambio
    
    Set oConect = New DConecta
    
    If oConect.AbreConexion = False Then Exit Function
    CON = PrnSet("C+")
    COFF = PrnSet("C-")
    BON = PrnSet("B+")
    BOFF = PrnSet("B-")
    
    sCabecera = ""
    sAsiento = ""
    nLi = 0
    If psMovNro = "" Then
       Exit Function
    End If
       lsFecha = GetFechaMov(psMovNro, True)
       sSQL = "SELECT * FROM Mov WHERE cMovNro = '" & psMovNro & "'"
       Set rs = oConect.CargaRecordSet(sSQL)
       If rs.EOF Then
          Exit Function
       End If
       lsGlosa = Replace(rs!cMovdesc, Chr(13) & oImpresora.gPrnSaltoLinea, " ")
       lsMovNro = psMovNro
       lsOpeCod = rs!cOpeCod
       If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Or Left(lsOpeCod, 4) = Left(gOpeME, 4) Or rs!cOpeCod = gContAjusteTipoCambio Then
        'TORE 20190801 - RFC1712260009
         sqlTC = "select CONVERT(date, mcr.dDocFecha) dDocFechaEmision   from Mov m inner join MovRef mr on m.nMovNro = mr.nMovNro inner join MovComprobanteReg mcr on mr.nMovNroRef = mcr.nMovNro where m.cMovNro = '" & psMovNro & "'"
         Set rsTC = oConect.CargaRecordSet(sqlTC)
         If Not rsTC.EOF Then
            lsTCFechaDoc = rsTC!dDocFechaEmision
         Else
            lsTCFechaDoc = lsFecha
         End If
         lnCambioVentaSBS = oTipoCamb.EmiteTipoCambio(CDate(Format(lsTCFechaDoc, "dd/MM/yyyy")), TCPondVenta)
         'END TORE 20190801 - RFC1712260009
         
         lnCambioFijo = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCFijoMes)
         lnCambioVenta = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCVenta)
         lnCambioCompra = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCCompra)
         'lnCambioVentaSBS = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCPondVenta)
         lnCambioCompraSBS = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCPonderado)
       End If
       
       If rs!nMovFlag = gMovFlagDeExtorno Then
            psTit1 = " A S I E N T O  C O N T A B L E   D E   E X T O R N O"
       ElseIf rs!nMovFlag = gMovFlagEliminado Then
            psTit1 = " **** ASIENTO ELIMINADO NO VALIDO PARA CONTABILIDAD **** "
       End If
       'ddd
       If Left(lsOpeCod, 3) = "622" Or Left(lsOpeCod, 3) = "632" Then
            sAsiento = sAsiento & ImpreCabAsiento(pnColPage, CDate(lsFecha), gsNomCmac, lsOpeCod, lsMovNro, psTit1, False, , , nLi) & oImpresora.gPrnSaltoLinea
       Else
            sAsiento = sAsiento & ImpreCabAsiento(pnColPage, CDate(lsFecha), gsNomCmac, lsOpeCod, lsMovNro, psTit1, True, , , nLi) & oImpresora.gPrnSaltoLinea
       End If
       If psCabe1 <> "" Then
          sAsiento = sAsiento & psCabe1 & oImpresora.gPrnSaltoLinea
          nLi = nLi + 4
       Else
           'BUSCA LAS PERSONA INVOLUCRADAS EN EL MOVIMIENTO
            sSQL = " SELECT  DISTINCT M.CMOVNRO,  " _
                    & "         CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                    & "                 WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                    & "         P.cPersCod, " _
                    & "         Convert(Char(40),P.CPERSNOMBRE) as cPersNombre, " _
                    & "         MOI.CPersCod ,isnull(PR.cPersIDNro,'') RUC,isnull(PR1.cPersIDNro,'') DNI" _
                    & "         ,isnull(PR2.cPersIDNro,'') Carnet" _
                    & " FROM    MOV M JOIN MOVCTA MC    ON M.NMOVNRO = MC.NMOVNRO " _
                    & "         JOIN MOVGasto MOI       ON MOI.NMOVNRO = M.NMOVNRO " _
                    & "         JOIN PERSONA P          ON P.CPERSCOD = MOI.CPERSCOD " _
                    & "         LEFT JOIN PERSID Pr          ON PR.CPERSCOD = P.CPERSCOD AND PR.cPersIdTpo =2" _
                    & "         LEFT JOIN PERSID Pr1         ON PR1.CPERSCOD = P.CPERSCOD AND PR1.cPersIdTpo =1" _
                    & "         LEFT JOIN PERSID Pr2         ON PR2.CPERSCOD = P.CPERSCOD AND PR2.cPersIdTpo =4" _
                    & " WHERE   M.cMovNro = '" & lsMovNro & "' "
              
            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = ""
            
            sDNI = IIf(rs!DNI = "", "", rs!DNI)
            sRUC = IIf(rs!RUC = "", "", rs!RUC)
            sCarnet = IIf(rs!Carnet = "", "", rs!Carnet)
            Do While Not rs.EOF
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & PstaNombre(rs!cpersNombre)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & PstaNombre(rs!cpersNombre)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            
            If sOrigen <> "" Or sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat("Persona(s) Origen  ", 20) & " :" & IIf(sOrigen = "", sDestino, sOrigen) & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
'            If sRUC <> "" Or sDNI <> "" Then
'                sAsiento = sAsiento & ImpreFormat(IIf(sRUC = "", "DNI", "RUC"), 20) & " :" & IIf(sRUC = "", sDNI, sRUC) & oImpresora.gPrnSaltoLinea
'                nLi = nLi + 1
'            End If

            If sRUC <> "" Then
               sAsiento = sAsiento & ImpreFormat("RUC", 20) & " :" & sRUC & oImpresora.gPrnSaltoLinea
            Else
                If sDNI <> "" Then
                   sAsiento = sAsiento & ImpreFormat(IIf(sRUC = "", "DNI", "RUC"), 20) & " :" & IIf(sRUC = "", sDNI, sRUC) & oImpresora.gPrnSaltoLinea
                Else
                   'Carnet de Extranjeria
                   sAsiento = sAsiento & ImpreFormat("CARNET", 20) & " :" & sCarnet & oImpresora.gPrnSaltoLinea
                End If
            End If
            nLi = nLi + 1
            
            
            
            sSQL = " SELECT  DISTINCT M.CMOVNRO, " _
                & "         CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                 WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "         P.cPersCod , CI.cCtaIFCod , " _
                & "         Convert(Char(40),CTPOI.cConsDescripcion ) as Entidad, " _
                & "         Convert(Char(40),P.CPERSNOMBRE) as cPersNombre, " _
                & "         Convert(char(50),C.cConsDescripcion + ' ' + CI.cCtaIFDesc) as DescCtaIF, " _
                & "         MO.COBJETOCOD " _
                & " FROM    MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "         JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "         JOIN MOVOBJIF MOI       ON MOI.NMOVNRO = MO.NMOVNRO AND MOI.NMOVITEM = MO.NMOVITEM " _
                & "         JOIN CTAIF CI           ON CI.CCTAIFCOD = MOI.CCTAIFCOD AND MOI.CPERSCOD = CI.CPERSCOD AND MOI.CIFTPO = CI.CIFTPO  " _
                & "         JOIN CONSTANTE C        ON C.nConsValor = SUBSTRING(CI.CCTAIFCOD,2,1) AND C.nCONSCOD LIKE '" & gCGTipoCtaIF & "' " _
                & "         JOIN INSTITUCIONFINANC I    ON I.CPERSCOD = CI.CPERSCOD AND I.CIFTPO = CI.CIFTPO " _
                & "         JOIN PERSONA P          ON P.CPERSCOD = I.CPERSCOD " _
                & "         JOIN CONSTANTE CTPOI        ON CTPOI.nConsValor = I.cIFTpo AND CTPOI.nCONSCOD LIKE '" & gCGTipoIF & "' " _
                & " WHERE M.cMovNro = '" & lsMovNro & "' AND MO.cObjetoCod ='" & Format(ObjEntidadesFinancieras, "00") & "'"
            
            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cpersNombre) & " " & Trim(rs!DescCtaIF)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cpersNombre) & " " & Trim(rs!DescCtaIF)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat("Cuenta(s) Origen", 20) & " :" + PrnSet("C+") + sOrigen + PrnSet("C-") + oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat("Cuenta(s) Destino", 20) & " :" + PrnSet("C+") + sDestino + PrnSet("C-") + oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
            sSQL = "    SELECT  DISTINCT M.CMOVNRO, " _
                & "             CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                     WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "             MOA.cAreaCod + MOA.CAGECOD AS cAreaCod, A.cAreaDescripcion + ' ' + ISNULL(AG.CAGEDESCRIPCION,'' )  as cAreaDescripcion , MO.COBJETOCOD, O.cObjetoDesc " _
                & "     FROM    MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "             JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "             JOIN MovObjAreaAgencia MOA         ON MOA.NMOVNRO = MO.NMOVNRO AND MOA.NMOVITEM = MO.NMOVITEM " _
                & "             JOIN AREAS A            ON A.cAreaCod = MOA.cAreaCod LEFT JOIN AGENCIAS AG ON AG.CAGECOD = MOA.CAGECOD " _
                & "             JOIN OBJETO O ON O.cObjetoCod  = MO.cObjetoCod " _
                & "     WHERE   M.cMovNro = '" & lsMovNro & "' AND Mo.cObjetoCod IN ('" & Format(ObjCMACAgenciaArea, "00") & "','" & Format(ObjCMACArea, "00") & "')"

            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                lsDesObjeto = Trim(rs!cObjetoDesc)
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!cAreaDescripcion)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!cAreaDescripcion)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Origen ", 20) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Destino ", 20) & " :" & sDestino & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
            sSQL = " SELECT    DISTINCT M.CMOVNRO, " _
                & "             CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                     WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "             MO.COBJETOCOD, O.cObjetoDesc " _
                & "   FROM      MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "             JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "             JOIN OBJETO O ON O.cObjetoCod  = MO.cObjetoCod " _
                & "   WHERE     M.cMovNro='" & lsMovNro & "' " _
                & "             AND MO.cObjetoCod NOT IN ('" & Format(ObjCMACAgencias, "00") & "','" & Format(ObjCMACAgenciaArea, "00") & "','" & Format(ObjDescomEfectivo, "00") & "','" & Format(ObjEntidadesFinancieras, "00") & "','" & Format(ObjBienesServicios, "00") & "','" & Format(ObjPersona, "00") & "')"
            
            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                lsDesObjeto = "OBJETOS"
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cObjetoCod) & "-" & Trim(rs!cObjetoDesc)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cObjetoCod) & "-" & Trim(rs!cObjetoDesc)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Origen ", 20) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Destino", 20) & " :" & sDestino & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
        End If ' del pscab1
    
       'Tipo de Cambio
        sSQL = "SELECT nMovTpoCambio, m.cOpeCod FROM MovTpoCambio MC JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE M.cMovNro = '" & lsMovNro & "'"
        Set rs = oConect.CargaRecordSet(sSQL)
        If Not rs.EOF Then
            If rs!cOpeCod = gContAjusteTipoCambio Then
            Else
               sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "       T.C.C SBS : " & Format(lnCambioCompraSBS, "##,###,##0.000") & "     T.C.V SBS : " & Format(lnCambioVentaSBS, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
               'sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "      T.C.Mercado: " & Format(rs!nMovTpoCambio, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
            End If
            nLi = nLi + 1
        Else
            If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Then
                sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "     T.C.Compra : " & Format(lnCambioCompra, "##,###,##0.000") & "     T.C.Venta : " & Format(lnCambioVenta, "##,###,##0.000") & "     T.C.C SBS : " & Format(lnCambioCompraSBS, "##,###,##0.000") & "     T.C.V SBS : " & Format(lnCambioVentaSBS, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
        End If
       nLi = nLi + 1
       'Datos de Documentos
       sSQL = "SELECT b.cDocAbrev, a.cDocNro, a.dDocFecha FROM MovDoc a join mov m on m.nmovnro =a.nmovnro  , " & vsBaseComunes & "Documento b " _
            & "WHERE m.cMovNro = '" & lsMovNro & "' and b.nDocTpo = a.nDocTpo "
       
        Set rsDoc = oConect.CargaRecordSet(sSQL)
        If Not rsDoc.EOF Then
          sDoc = ""
          nLi = 1
          Do While Not rsDoc.EOF
             sDoc = sDoc & rsDoc!cDocAbrev & "_" & rsDoc!cDocNro & " " & rsDoc!dDocFecha & "  "
             If nLi Mod 4 = 0 Then
                sDoc = sDoc + oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
             End If
             rsDoc.MoveNext
          Loop
          sAsiento = sAsiento + ImpreFormat("Documentos", 20) & " :" + PrnSet("C+") + sDoc + PrnSet("C-") + oImpresora.gPrnSaltoLinea
          nLi = nLi + 1
        End If
        rsDoc.Close
       'Glosa
       sAsiento = sAsiento & ImpreGlosa(lsGlosa, pnColPage, ImpreFormat("Glosa       ", 20) & " :", , , nLi) + oImpresora.gPrnSaltoLinea
       nLi = nLi + 1
       lsMovBilletaje = ""
       lsMovBilletaje = ImprimeBilletaje(Mid(lsOpeCod, 3, 1), lsMovNro, pnColPage, , nLi)
       If lsMovBilletaje <> "" Then
            sAsiento = sAsiento + lsMovBilletaje
       End If
    'Cabecera 2 de Asiento
    sCabecera = sCabecera + CON
    sCabecera = sCabecera + "====================================================================================================================================" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "Itm C U E N T A   C O N T A B L E                                                                           DEBE            HABER" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "    Codigo              Descripcion" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "------------------------------------------------------------------------------------------------------------------------------------" & COFF & oImpresora.gPrnSaltoLinea
    'MovCta para MN
    sSQL = "SELECT mc.nMovItem, mc.cCtaContCod, ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),'') cCtaContDesc, mc.nMovImporte " _
         & "FROM   MovCta mc JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE  M.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
    Set rs = oConect.CargaRecordSet(sSQL)
    If rs.EOF Then
       'MsgBox "No existen datos de Movimiento ", vbInformation, "Error"
       Exit Function
    End If
    nTot = 0
    nTotH = 0
    sAsiento = sAsiento + sCabecera
    nLi = nLi + 4
    sAsiento = sAsiento + BON & " EN MONEDA NACIONAL " + oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento + "--------------------" & BOFF + oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento & oImpresora.gPrnSaltoLinea
    
    sAsiento8 = "": sAsiento8c = ""
    Do While Not rs.EOF
       DoEvents
       If nLi > pnLinPage - 6 Then
           nLi = 5
           sAsiento = sAsiento + oImpresora.gPrnSaltoPagina & sCabecera
       End If
       
       If Left(rs!cCtaContCod, 1) = "8" Then '*** PEAC 20110325
            sAsiento8 = sAsiento8 + CON & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " & Justifica(rs!cCtaContDesc, 71) & " " _
            & Right(Space(16) & IIf(rs!nMovImporte > 0, Format(rs!nMovImporte, "##,###,##0.00"), ""), 16) & " " _
            & Right(Space(16) & IIf(rs!nMovImporte < 0, Format(rs!nMovImporte * -1, "##,###,##0.00"), ""), 16) _
            & COFF & oImpresora.gPrnSaltoLinea
            
            If rs!nMovImporte > 0 Then
            nTot8 = nTot8 + Val(rs!nMovImporte)
            Else
            nTotH8 = nTotH8 + Val(rs!nMovImporte) * -1
            End If
            rs.MoveNext
       Else
            sAsiento = sAsiento + CON & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " & Justifica(rs!cCtaContDesc, 71) & " " _
            & Right(Space(16) & IIf(rs!nMovImporte > 0, Format(rs!nMovImporte, "##,###,##0.00"), ""), 16) & " " _
            & Right(Space(16) & IIf(rs!nMovImporte < 0, Format(rs!nMovImporte * -1, "##,###,##0.00"), ""), 16) _
            & COFF & oImpresora.gPrnSaltoLinea
            
            
            If rs!nMovImporte > 0 Then
            nTot = nTot + Val(rs!nMovImporte)
            Else
            nTotH = nTotH + Val(rs!nMovImporte) * -1
            End If
            nLi = nLi + 1
            rs.MoveNext
       End If
    Loop
    
    If sAsiento8 <> "" Then '*** PEAC 20110425
        sAsiento8c = oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        sAsiento8c = sAsiento8c + "A S I E N T O   D E   C U E N T A   D E   O R D E N" & oImpresora.gPrnSaltoLinea
        sAsiento8c = sAsiento8c + String(pnColPage - 1, "=") & oImpresora.gPrnSaltoLinea
        sAsiento8c = sAsiento8c + CON + "Cuenta Contable                                           DEBE       HABER" & oImpresora.gPrnSaltoLinea
        sAsiento8c = sAsiento8c + String(pnColPage - 1, "-") & oImpresora.gPrnSaltoLinea
        sAsiento8c = sAsiento8c + sAsiento8
        sAsiento8c = sAsiento8c + String(pnColPage - 1, "-") & oImpresora.gPrnSaltoLinea
        '''sAsiento8c = sAsiento8c + Space(75) & "TOTALES        " & gcMN & Right(Space(15) & Format(nTot8, "###,###,##0.00"), 15) & "  " & gcMN & Right(Space(15) & Format(nTotH8, "###,###,##0.00"), 15) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
        sAsiento8c = sAsiento8c + Space(75) & "TOTALES        " & gcPEN_SIMBOLO & Right(Space(15) & Format(nTot8, "###,###,##0.00"), 15) & "  " & gcPEN_SIMBOLO & Right(Space(15) & Format(nTotH8, "###,###,##0.00"), 15) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
        sAsiento8c = sAsiento8c + String(pnColPage - 1, "=") & oImpresora.gPrnSaltoLinea
    End If
    
    sAsiento = sAsiento + CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF & oImpresora.gPrnSaltoLinea
    
    sAsiento = sAsiento + sAsiento8c '*** PEAC 20110425
    
    nLi = nLi + 2

    sSQL = "SELECT mc.nMovItem, mc.cCtaContCod, LTRIM(RTRIM(ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),''))) cCtaContDesc, me.nmovMEimporte " _
         & "FROM   mov m join MovCta mc on mc.nmovnro=m.nmovnro JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
         & "WHERE  m.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
    Set rs = oConect.CargaRecordSet(sSQL)
    sTexto = ""
    nTot = 0
    nTotH = 0
    If Not rs.EOF Then
       sAsiento = sAsiento + BON & " EN MONEDA EXTRANJERA " & oImpresora.gPrnSaltoLinea
       sAsiento = sAsiento + "----------------------" & BOFF & oImpresora.gPrnSaltoLinea
       nLi = nLi + 2
    End If
    Do While Not rs.EOF
       DoEvents
       If nLi > pnLinPage - 6 Then
          nLi = 5
          sTexto = sTexto + oImpresora.gPrnSaltoPagina
          sTexto = sTexto + sCabecera
       End If
       sTexto = sTexto + CON & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " _
              & Justifica(rs!cCtaContDesc, 71) & " " _
              & IIf(rs!nMovMEImporte > 0, PrnVal(rs!nMovMEImporte, 16, 2), Space(16)) & " " _
              & IIf(rs!nMovMEImporte < 0, PrnVal(rs!nMovMEImporte * -1, 16, 2), Space(16)) _
              & COFF & oImpresora.gPrnSaltoLinea
       If rs!nMovMEImporte > 0 Then
          nTot = nTot + Val(rs!nMovMEImporte)
       Else
          nTotH = nTotH + Val(rs!nMovMEImporte) * -1
       End If
       nLi = nLi + 1
       rs.MoveNext
    Loop
    If nTot > 0 Or nTotH > 0 Then
       sAsiento = sAsiento & sTexto
       sAsiento = sAsiento + CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
       sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF + oImpresora.gPrnSaltoLinea
       nLi = nLi + 3
    End If
    If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
    
    If lsOpeCod <> 421113 Then
        sAsiento = sAsiento + CON & "====================================================================================================================================" & COFF & oImpresora.gPrnSaltoLinea
        
        sAsiento = sAsiento + ImprePiePag(pnColPage, psPiePag, psOtraFirma)
    End If
    sAsiento = sAsiento + "   " + oImpresora.gPrnSaltoLinea

    Set oDGeneral = Nothing
    Set oConect = Nothing
    ImprimeAsientoContableProv = sAsiento
    
    Exit Function
ImprimeAsientoContableErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeAsientoContable Method")
End Function

Public Function RepInstPubliResgos(Optional pdFechas As Date = "1990-01-01") As String
    On Error GoTo RepInstPubliResgosErr
    Dim sSQL As String
    Dim sSQL1 As String
    Dim sCabecera As String
    Dim sAsiento As String
    Dim nLi As Long
            
    Dim rs       As New ADODB.Recordset
    Dim rs1       As New ADODB.Recordset
    Dim oConect As DConecta
    Set oConect = New DConecta
    Dim oDBalanceContA As DbalanceCont
    Set oDBalanceContA = New DbalanceCont
    Dim nSSoles As Currency
    Dim nSDolares As Currency
    If oConect.AbreConexion = False Then Exit Function
    
    sCabecera = ""
    sAsiento = ""
    nLi = 0
    
    'Sacamos el Saldo de Cuentas de Ahorros  en soles y Dolares
    'Ahorros
    '***Modificado por ELRO el 20111104, según Acta 306-2011/TI-D
    '     sSql = " Select substring(pr.cCtaCod,9,1)Moneda,isnull(Sum(pr.nSaldo),0)Saldo" & _
    '            " From Producto pr " & _
    '            " Inner join  ProductoPersona pp ON pp.cCtaCod=pr.cCtaCod " & _
    '            " Inner join Persona p ON p.cPersCod =pp.cPersCod " & _
    '            " Where pr.nPrdEstado not in ('100300','100400')and substring(pp.cctaCod,6,3)='232'" & _
    '            " and substring(pr.cCtaCod,9,1)in (1) and p.cPersCod in (Select cPErsCod FRom PersonasGob) " & _
    '            " Group By substring(pr.cCtaCod,9,1) " & _
    '            " Union All " & _
    '            " Select substring(pr.cCtaCod,9,1)Moneda,isnull(Sum(pr.nSaldo),0)Saldo" & _
    '            " From Producto pr " & _
    '            " Inner join  ProductoPersona pp ON pp.cCtaCod=pr.cCtaCod " & _
    '            " Inner join Persona p ON p.cPersCod =pp.cPersCod " & _
    '            " Where pr.nPrdEstado not in ('100300','100400')and substring(pp.cctaCod,6,3)='232' " & _
    '            " and substring(pr.cCtaCod,9,1)in (2) and p.cPersCod in (Select cPErsCod FRom PersonasGob)" & _
    '            " Group By substring(pr.cCtaCod,9,1) "
         sSQL = " Select substring(pr.cCtaCod,9,1)Moneda,isnull(Sum(pr.nSaldo),0)Saldo" & _
                " From Producto pr " & _
                " Inner join  ProductoPersona pp ON pp.cCtaCod=pr.cCtaCod " & _
                " Inner join Persona p ON p.cPersCod =pp.cPersCod " & _
                " Where pr.nPrdEstado not in ('1300','1400')and substring(pp.cctaCod,6,3)='232'" & _
                " and substring(pr.cCtaCod,9,1)in (1) and p.cPersCod in (Select cPErsCod FRom PersonasGob) " & _
                " Group By substring(pr.cCtaCod,9,1) " & _
                " Union All " & _
                " Select substring(pr.cCtaCod,9,1)Moneda,isnull(Sum(pr.nSaldo),0)Saldo" & _
                " From Producto pr " & _
                " Inner join  ProductoPersona pp ON pp.cCtaCod=pr.cCtaCod " & _
                " Inner join Persona p ON p.cPersCod =pp.cPersCod " & _
                " Where pr.nPrdEstado not in ('1300','1400')and substring(pp.cctaCod,6,3)='232' " & _
                " and substring(pr.cCtaCod,9,1)in (2) and p.cPersCod in (Select cPErsCod FRom PersonasGob)" & _
                " Group By substring(pr.cCtaCod,9,1) "
    '***Fin Modificado por ELRO**********************************
    
       Set rs = oConect.CargaRecordSet(sSQL)

    'Sacamos el Saldo de Cuentas de Plazo Fijo en soles y Dolares
    'Plazo Fijo
    '***Modificado por ELRO el 20111104, según Acta 306-2011/TI-D
    '     sSQL1 = " Select substring(pr.cCtaCod,9,1)Moneda,isnull(Sum(pr.nSaldo),0)Saldo" & _
    '            " From Producto pr " & _
    '            " Inner join  ProductoPersona pp ON pp.cCtaCod=pr.cCtaCod " & _
    '            " Inner join Persona p ON p.cPersCod =pp.cPersCod " & _
    '            " Where pr.nPrdEstado not in ('100300','100400')and substring(pp.cctaCod,6,3)='233'" & _
    '            " and substring(pr.cCtaCod,9,1)in (1) and p.cPersCod in (Select cPErsCod FRom PersonasGob) " & _
    '            " Group By substring(pr.cCtaCod,9,1) " & _
    '            " Union All " & _
    '            " Select substring(pr.cCtaCod,9,1)Moneda,isnull(Sum(pr.nSaldo),0)Saldo " & _
    '            " From Producto pr " & _
    '            " Inner join  ProductoPersona pp ON pp.cCtaCod=pr.cCtaCod " & _
    '            " Inner join Persona p ON p.cPersCod =pp.cPersCod " & _
    '            " Where pr.nPrdEstado not in ('100300','100400')and substring(pp.cctaCod,6,3)='233' " & _
    '            " and substring(pr.cCtaCod,9,1)in (2) and p.cPersCod in (Select cPErsCod FRom PersonasGob) " & _
    '            " Group By substring(pr.cCtaCod,9,1) "
    sSQL1 = " Select substring(pr.cCtaCod,9,1)Moneda,isnull(Sum(pr.nSaldo),0)Saldo" & _
            " From Producto pr " & _
            " Inner join  ProductoPersona pp ON pp.cCtaCod=pr.cCtaCod " & _
            " Inner join Persona p ON p.cPersCod =pp.cPersCod " & _
            " Where pr.nPrdEstado not in ('1300','1400')and substring(pp.cctaCod,6,3)='233'" & _
            " and substring(pr.cCtaCod,9,1)in (1) and p.cPersCod in (Select cPErsCod FRom PersonasGob) " & _
            " Group By substring(pr.cCtaCod,9,1) " & _
            " Union All " & _
            " Select substring(pr.cCtaCod,9,1)Moneda,isnull(Sum(pr.nSaldo),0)Saldo " & _
            " From Producto pr " & _
            " Inner join  ProductoPersona pp ON pp.cCtaCod=pr.cCtaCod " & _
            " Inner join Persona p ON p.cPersCod =pp.cPersCod " & _
            " Where pr.nPrdEstado not in ('1300','1400')and substring(pp.cctaCod,6,3)='233' " & _
            " and substring(pr.cCtaCod,9,1)in (2) and p.cPersCod in (Select cPErsCod FRom PersonasGob) " & _
            " Group By substring(pr.cCtaCod,9,1) "
    '***Fin Modificado por ELRO**********************************
    
       Set rs1 = oConect.CargaRecordSet(sSQL1)
        
       If rs.EOF And rs1.EOF Then
          Exit Function
       End If
        
        
    ' cabecera
    sCabecera = sCabecera & "SALDO DE DEPOSITOS DE GRANDES ACREEDORES DEL ANEXO 15-A-CUENTAS DEL ESTADO" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera & Space(15) & "CUENTA DE AHORROS" & oImpresora.gPrnSaltoLinea
    '''sCabecera = sCabecera & Space(5) & "S/." & Space(20) & "$." & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    sCabecera = sCabecera & Space(5) & gcPEN_SIMBOLO & Space(20) & "$." & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    sCabecera = sCabecera & String(40, "_") & oImpresora.gPrnSaltoLinea
    
    Dim MontoSoles As Currency
    Dim MontoDolares As Currency
    
    MontoSoles = 0
    MontoDolares = 0
    
    Do While Not rs.EOF
       DoEvents
       If rs!Moneda = 1 Then
          MontoSoles = rs!Saldo
          nSSoles = nSSoles + MontoSoles
       Else
          MontoDolares = rs!Saldo
          nSDolares = nSDolares + MontoDolares
       End If
       rs.MoveNext
    Loop
    sAsiento = sAsiento & Space(4) & Format(MontoSoles, "0.00") & Space(20 - Len(MontoSoles)) & Format(MontoDolares, "0.00") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    
    sAsiento = sAsiento & Space(15) & "DEPOSITO A PLAZO FIJO" & oImpresora.gPrnSaltoLinea
    '''sAsiento = sAsiento & Space(5) & "S/." & Space(20) & "$." & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    sAsiento = sAsiento & Space(5) & gcPEN_SIMBOLO & Space(20) & "$." & oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento & String(40, "_") & oImpresora.gPrnSaltoLinea
'
'    MontoSoles = 0
'    MontoDolares = 0
    
    Do While Not rs1.EOF
       DoEvents
       If rs1!Moneda = 1 Then
          MontoSoles = rs1!Saldo
          nSSoles = nSSoles + MontoSoles
       Else
          MontoDolares = rs1!Saldo
          nSDolares = nSDolares + MontoDolares
       End If
       rs1.MoveNext
    Loop
    sAsiento = sAsiento & Space(4) & Format(MontoSoles, "0.00") & Space(20 - Len(MontoSoles)) & Format(MontoDolares, "0.00") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    Call oDBalanceContA.InsertaCtaContSaldoDiario("Estado_1", pdFechas, "763406", nSSoles)
    Call oDBalanceContA.InsertaCtaContSaldoDiario("Estado_2", pdFechas, "763406", nSDolares)
    sAsiento = sCabecera & sAsiento
    RepInstPubliResgos = sAsiento
    Exit Function
RepInstPubliResgosErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:RepInstPubliResgos Method")
End Function

'*** PEAC 20091113
Public Function ImprimirFactura(ByVal pMatDetalle As Variant, _
                                       ByVal pMatDatos As Variant, ByVal pnNumFilas As Integer, ByVal pnMoneda As Integer, ByVal pnTipoDescri As Integer) As String

Dim sTexto As String
Dim nAncho As Integer, N As Integer
Dim lsGlosa As String
Dim I As Integer
Dim lnImporte As Currency
Dim lsMovNro As String
Dim lnTotal As Currency
Dim lsTitulo As String
Dim sDetTexto As String
sTexto = ""
sDetTexto = ""
Dim lcNomMes As String
Dim lnMone As Integer
'lcNomMes = ArmaFecha(CDate(pMatDatos(3)))
lcNomMes = Choose(Month(pMatDatos(3)), "ENERO", "FEBRERO", "MARZO", "ABRIL", _
                                        "MAYO", "JUNIO", "JULIO", "AGOSTO", _
                                        "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE")

sTexto = sTexto & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
sTexto = sTexto & Space(65) & pMatDatos(1) & " - " & pMatDatos(2) & oImpresora.gPrnSaltoLinea
sTexto = sTexto & Space(10) & pMatDatos(4) & oImpresora.gPrnSaltoLinea
sTexto = sTexto & Space(10) & pMatDatos(5) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
sTexto = sTexto & Space(10) & Trim(pMatDatos(6)) & Space(Len(Space(57)) - Len(Trim(pMatDatos(6)))) & Format(Day(pMatDatos(3)), "00") & Space(5) & Justifica(lcNomMes, 10) & Space(7) & Trim(Str(Year(pMatDatos(3)))) & oImpresora.gPrnSaltoLinea
sTexto = sTexto & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea

    '************************************************************
    For I = 0 To pnNumFilas - 2
        If I = 0 And pMatDetalle(1, I) = 0 And pMatDetalle(3, I) = 0 And pMatDetalle(4, I) = 0 Then
            sTexto = sTexto & PrnVal(Str(pMatDetalle(1, I)), 12, 2, False) + Space(2) + Justifica(CStr(pMatDetalle(2, I)), 50) + Space(5) + PrnVal(Str(pMatDatos(7)), 12, 2, False) + Space(3) + PrnVal(Str(pMatDatos(7)), 12, 2, False)
        Else
            sTexto = sTexto & PrnVal(Str(pMatDetalle(1, I)), 12, 2, False) + Space(2) + Justifica(CStr(pMatDetalle(2, I)), 50) + Space(5) + PrnVal(Str(pMatDetalle(3, I)), 12, 2, False) + Space(3) + PrnVal(Str(pMatDetalle(4, I)), 12, 2, False)
        End If
        sTexto = sTexto & oImpresora.gPrnSaltoLinea
    Next I
    
    'imprime los espacios para mantener LA LINEA DIEZ
    If I < 10 Then
        For N = 1 To 10 - I
            sTexto = sTexto & oImpresora.gPrnSaltoLinea
        Next N
    End If
    sTexto = sTexto & oImpresora.gPrnSaltoLinea
    '************************************************************

If pnMoneda = 1 Then
    sTexto = sTexto & Space(14) & "SON: " & ConvNumLet(CStr(pMatDatos(9)), True, False, 1) & oImpresora.gPrnSaltoLinea
Else
    sTexto = sTexto & Space(14) & "SON: " & ConvNumLet(CStr(pMatDatos(9)), True, False, 2) & oImpresora.gPrnSaltoLinea
End If

sTexto = sTexto & oImpresora.gPrnSaltoLinea
'''sTexto = sTexto + Space(83) & IIf(pnMoneda = 1, "S/.", "US$") & Space(0) & PrnVal(Str(pMatDatos(7)), 10, 2, True) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
sTexto = sTexto + Space(83) & IIf(pnMoneda = 1, gcPEN_SIMBOLO, "US$") & Space(0) & PrnVal(Str(pMatDatos(7)), 10, 2, True) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
sTexto = sTexto + Space(84) & PrnVal(Str(pMatDatos(8)), 12, 2, True) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'''sTexto = sTexto + Space(83) & IIf(pnMoneda = 1, "S/.", "US$") & Space(0) & PrnVal(Str(pMatDatos(9)), 10, 2, True) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
sTexto = sTexto + Space(83) & IIf(pnMoneda = 1, gcPEN_SIMBOLO, "US$") & Space(0) & PrnVal(Str(pMatDatos(9)), 10, 2, True) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016

ImprimirFactura = sTexto

End Function

'*** PEAC 20091113
Public Function ImprimirCheque(ByVal pMatDatos As Variant) As String

Dim sTexto As String
Dim nAncho As Integer, N As Integer
Dim lsGlosa As String
Dim I As Integer
Dim lnImporte As Currency
Dim lsMovNro As String
Dim lnTotal As Currency
Dim lsTitulo As String

Dim lcNomMes As String
Dim lnMone As Integer
Dim cBco1, cBco2, cBco3, cBco4, cBco5, cBco6 As String

sTexto = " "

cBco1 = "1090100824682" ''continen
cBco2 = "1090100824674" ''interbank
cBco3 = "1090100824640" ''credito
cBco4 = "1090100835102" ''scotia
cBco5 = "1090100824631" ''bco de la nacion
cBco6 = "1090100012521" ''caja maynas

Select Case pMatDatos(4)
    Case 1 '' soles
        Select Case pMatDatos(5)
            Case cBco1  '' continental sol
                Linea sTexto, BON & Space(34) & Format(Mid(pMatDatos(1), 1, 2), "00") & Space(2) & Format(Mid(pMatDatos(1), 4, 2), "00") & Space(2) & Format(Mid(pMatDatos(1), 9, 2), "00") & Space(8) & "**" & Format(pMatDatos(3), "#,#.00") & "**" & BOFF, 4
                Linea sTexto, BON & Space(9) & pMatDatos(2) & BOFF, 2
                Linea sTexto, oImpresora.gPrnTamLetra15CPI & BON & Space(2) & "**" & Trim(ConvNumLet(CStr(pMatDatos(3)), True)) & "**" & Chr$(27) & Chr$(77) & BOFF, 2
            Case cBco2  '' interbank sol
                Linea sTexto, BON & Space(30) & Format(Mid(pMatDatos(1), 1, 2), "00") & Space(1) & Format(Mid(pMatDatos(1), 4, 2), "00") & Space(1) & Format(Mid(pMatDatos(1), 9, 2), "00") & BOFF, 1
                Linea sTexto, BON & Space(49) & "**" & Format(pMatDatos(3), "#,#.00") & "**" & BOFF, 3
                Linea sTexto, BON & Space(9) & pMatDatos(2) & BOFF, 2
                Linea sTexto, oImpresora.gPrnTamLetra15CPI & BON & Space(0) & "**" & Trim(ConvNumLet(CStr(pMatDatos(3)), True)) & "**" & Chr$(27) & Chr$(77) & BOFF, 2
            Case cBco3  '' credito sol
                Linea sTexto, BON & Space(38) & Format(Mid(pMatDatos(1), 1, 2), "00") & Space(1) & Format(Mid(pMatDatos(1), 4, 2), "00") & Space(1) & Format(Mid(pMatDatos(1), 9, 2), "00") & Space(7) & "**" & Format(pMatDatos(3), "#,#.00") & "**" & BOFF, 3
                Linea sTexto, BON & Space(9) & pMatDatos(2) & BOFF, 2
                Linea sTexto, oImpresora.gPrnTamLetra15CPI & BON & Space(0) & "**" & Trim(ConvNumLet(CStr(pMatDatos(3)), True)) & "**" & Chr$(27) & Chr$(77) & BOFF, 2
            Case cBco4  ''scotia sol
                Linea sTexto, BON & Space(29) & Format(Mid(pMatDatos(1), 1, 2), "00") & Space(1) & Format(Mid(pMatDatos(1), 4, 2), "00") & Space(2) & Format(Mid(pMatDatos(1), 9, 2), "00") & Space(15) & "**" & Format(pMatDatos(3), "#,#.00") & "**" & BOFF, 4
                Linea sTexto, BON & Space(10) & pMatDatos(2) & BOFF, 2
                Linea sTexto, oImpresora.gPrnTamLetra15CPI & BON & Space(8) & "**" & Trim(ConvNumLet(CStr(pMatDatos(3)), True)) & "**" & Chr$(27) & Chr$(77) & BOFF, 2
            Case cBco5   '' nacion sol
                Linea sTexto, oImpresora.gPrnTamLetra15CPI & BON & Space(35) & Trim(pMatDatos(6)) & Space(Len(Space(20)) - Len(Trim(pMatDatos(6)))) & Format(Mid(pMatDatos(1), 1, 2), "00") & Space(4) & Format(Mid(pMatDatos(1), 4, 2), "00") & Space(4) & Format(Mid(pMatDatos(1), 9, 2), "00") & Space(20) & "**" & Format(pMatDatos(3), "#,#.00") & "**" & BOFF, 3
                Linea sTexto, BON & Space(10) & pMatDatos(2) & BOFF, 2
                Linea sTexto, oImpresora.gPrnTamLetra15CPI & BON & Space(4) & "**" & Trim(ConvNumLet(CStr(pMatDatos(3)), True)) & "**" & Chr$(27) & Chr$(77) & BOFF, 2
            Case cBco6 '' caja maynas sol
                Linea sTexto, BON & Space(29) & Format(Mid(pMatDatos(1), 1, 2), "00") & Space(3) & Format(Mid(pMatDatos(1), 4, 2), "00") & Space(3) & Format(Mid(pMatDatos(1), 9, 2), "00") & Space(11) & "**" & Format(pMatDatos(3), "#,#.00") & "**" & BOFF, 4
                Linea sTexto, BON & Space(11) & pMatDatos(2) & BOFF, 3
                Linea sTexto, oImpresora.gPrnTamLetra15CPI & BON & Space(0) & "**" & Trim(ConvNumLet(CStr(pMatDatos(3)), True)) & "**" & Chr$(27) & Chr$(77) & BOFF, 2
        End Select
    Case 2 '' dolar
        Select Case pMatDatos(5)
            Case cBco1 '' contine dol
                Linea sTexto, BON & Space(34) & Format(Mid(pMatDatos(1), 1, 2), "00") & Space(2) & Format(Mid(pMatDatos(1), 4, 2), "00") & Space(2) & Format(Mid(pMatDatos(1), 9, 2), "00") & Space(8) & "**" & Format(pMatDatos(3), "#,#.00") & "**" & BOFF, 4
                Linea sTexto, BON & Space(9) & pMatDatos(2) & BOFF, 2
                Linea sTexto, oImpresora.gPrnTamLetra15CPI & BON & Space(2) & "**" & Trim(ConvNumLet(CStr(pMatDatos(3)), True)) & "**" & Chr$(27) & Chr$(77) & BOFF, 2
            Case cBco2 '' interba dol
                Linea sTexto, BON & Space(33) & Format(Mid(pMatDatos(1), 1, 2), "00") & Space(2) & Format(Mid(pMatDatos(1), 4, 2), "00") & Space(2) & Format(Mid(pMatDatos(1), 9, 2), "00") & Space(8) & "**" & Format(pMatDatos(3), "#,#.00") & "**" & BOFF, 4
                Linea sTexto, BON & Space(9) & pMatDatos(2) & BOFF, 2
                Linea sTexto, oImpresora.gPrnTamLetra15CPI & BON & Space(2) & "**" & Trim(ConvNumLet(CStr(pMatDatos(3)), True)) & "**" & Chr$(27) & Chr$(77) & BOFF, 2
            Case cBco3 '' credito dol
                Linea sTexto, BON & Space(38) & Format(Mid(pMatDatos(1), 1, 2), "00") & Space(1) & Format(Mid(pMatDatos(1), 4, 2), "00") & Space(1) & Format(Mid(pMatDatos(1), 9, 2), "00") & Space(7) & "**" & Format(pMatDatos(3), "#,#.00") & "**" & BOFF, 3
                Linea sTexto, BON & Space(9) & pMatDatos(2) & BOFF, 2
                Linea sTexto, oImpresora.gPrnTamLetra15CPI & BON & Space(0) & "**" & Trim(ConvNumLet(CStr(pMatDatos(3)), True)) & "**" & Chr$(27) & Chr$(77) & BOFF, 2
            Case cBco4  '' scotia dol
                Linea sTexto, BON & Space(29) & Format(Mid(pMatDatos(1), 1, 2), "00") & Space(1) & Format(Mid(pMatDatos(1), 4, 2), "00") & Space(2) & Format(Mid(pMatDatos(1), 9, 2), "00") & Space(15) & "**" & Format(pMatDatos(3), "#,#.00") & "**" & BOFF, 4
                Linea sTexto, BON & Space(10) & pMatDatos(2) & BOFF, 2
                Linea sTexto, oImpresora.gPrnTamLetra15CPI & BON & Space(8) & "**" & Trim(ConvNumLet(CStr(pMatDatos(3)), True)) & "**" & Chr$(27) & Chr$(77) & BOFF, 2
        End Select
End Select
ImprimirCheque = sTexto
End Function
Public Function ImprimeListadodeCalificacion(pdFecha As Date, _
                pnLinPage As Integer) As String
Dim nLin   As Integer, P As Integer
Dim lnFila As Integer
Dim sTexto As String
Dim sMonedaDes As String
Dim sEstadoDes As String
Dim sCondicDes As String
Dim sProducDes As String
Dim sAgeCod As String
Dim prs As ADODB.Recordset
Dim nSaldoCap As Currency
Dim nIntDeven As Currency
Dim nIntSuspe As Currency
Dim nProvisio As Currency
Dim nSaldoVen As Currency
Dim nCambio As Integer
Dim sProducDescripcion As String
Dim lnCont As Double
Dim lsImpre As String
nLin = 76
Dim dSaldo As DCtaSaldo
Set dSaldo = New DCtaSaldo
Set prs = dSaldo.ObtenerDatosInteresDevSuspen()
lnFila = 7
ReDim sTotales(3, 8, 4)
Do While Not prs.EOF
      DoEvents
      sTexto = sTexto & ImprimeListadoDeComprobacionCabecera(pdFecha, pdFecha, 0, P, 0, 1, pdFecha, nLin, pnLinPage)
   
     nCambio = 0
If sProducDes = prs!cProducto And sCondicDes = prs!cCalGen And sMonedaDes = prs!cMoneda And sEstadoDes = prs!nConsValor And sAgeCod = prs!cAgeCod Then
    nCambio = 0
Else
    nCambio = 1
End If
   If lnCont > 0 Then
       If nCambio = 1 Then
         Linea sTexto, BON & Space(25) & "Sub-Total Tipo " & Mid(prs!cConsDescripcion & Space(15), 1, 10) & "  ----->" & Space(27) & PrnVal(nSaldoCap, 14, 2) & PrnVal(nIntDeven, 14, 2) & PrnVal(nIntSuspe, 13, 2) & PrnVal(nProvisio, 14, 2) & PrnVal(nSaldoVen, 13, 2) & BOFF
            nSaldoCap = 0
            nIntDeven = 0
            nIntSuspe = 0
            nProvisio = 0
            nSaldoVen = 0
        End If
    End If
    If nCambio = 1 Then
        Linea sTexto, prs!cAgeDescripcion
        Linea sTexto, prs!cDescEstado
        Linea sTexto, IIf(prs!cMoneda = "1", "MONEDA NACIONAL", "MONEDA EXTRANEJERA")
        Linea sTexto, prs!cCondicion
        Linea sTexto, prs!cConsDescripcion
    End If
   sMonedaDes = prs!cMoneda
   sEstadoDes = prs!nConsValor
   sCondicDes = prs!cCalGen
   sProducDes = prs!cProducto
   sProducDescripcion = prs!cConsDescripcion
   sAgeCod = prs!cAgeCod
   Linea sTexto, Mid(prs!cConsDescripcion & Space(15), 1, 10) & Space(2) & prs!cCondicion & Space(12 - Len(Trim(prs!cCondicion))) & Mid(prs!cLineaDeta & Space(20), 7, 4) & Space(12) & Mid(prs!cLineaDeta & Space(100), 1, 33) & Space(1) & PrnVal(prs!nMontoApr, 14, 2) & PrnVal(prs!nSaldoCap, 14, 2) & PrnVal(prs!nIntDev, 14, 2) & PrnVal(prs!nIntSusp, 13, 2) & PrnVal(prs!nProvision + prs!nProvisionProciclica, 14, 2) & PrnVal(prs!nCapVencido, 13, 2)
   
    nSaldoCap = nSaldoCap + prs!nSaldoCap
    nIntDeven = nIntDeven + prs!nIntDev
    nIntSuspe = nIntSuspe + prs!nIntSusp
    nProvisio = nProvisio + prs!nProvision + prs!nProvisionProciclica
    nSaldoVen = nSaldoVen + prs!nCapVencido
      nLin = nLin + 1
      lnCont = lnCont + 1
      If lnCont Mod 600 = 0 Then
         lsImpre = lsImpre & sTexto
         sTexto = ""
      End If
   prs.MoveNext
Loop
     Linea sTexto, BON & Space(25) & "Sub-Total Tipo " & Mid(sProducDescripcion & Space(15), 1, 10) & "  ----->" & Space(27) & PrnVal(nSaldoCap, 16, 2) & PrnVal(nIntDeven, 14, 2) & PrnVal(nIntSuspe, 13, 2) & PrnVal(nProvisio, 14, 2) & PrnVal(nSaldoVen, 13, 2) & BOFF
        nSaldoCap = 0
        nIntDeven = 0
        nIntSuspe = 0
        nProvisio = 0
        nSaldoVen = 0

ImprimeListadodeCalificacion = lsImpre & sTexto
End Function

Private Function ImprimeListadoDeComprobacionCabecera(pdFechaIni As Date, pdFechaFin As Date, pnTpoCambio As Currency, P As Integer, pnTipoBala As Integer, pnMoneda As Integer, pdFecha As Date, ByRef nLin As Integer, ByRef pnLinPage As Integer) As String
Dim sTexto As String
If nLin > pnLinPage - 6 Then
   
   
   Linea sTexto, Cabecera(" LISTADO DE CALIFICACION DE CARTRA -INTERES DEVENGADOS Y EN SUSPENSO - PROVISIONES  ", P, "", 120, , Format(pdFecha, gsFormatoFechaView), lsEmprLogo), 0
   If pnMoneda = 0 Then
      Linea sTexto, Centra(" C O N S O L I D A D O ", 120)
   End If
   
   Linea sTexto, Centra("( DEL " & pdFechaIni & " AL " & pdFechaFin & ")", 120)
     
   Linea sTexto, " ================================================================================================================================================================="
   Linea sTexto, BON & "  TIPO    ¦ CONDICION  ¦ LINEA        ¦  DESCRIPCION                  ¦  MONTOAPRO    ¦ MONTSALDCAP ¦INTDEV       ¦INTESUSPENS ¦ PROVISION   ¦ SALDOCUOVENC" & BOFF
   Linea sTexto, " ================================================================================================================================================================="
   
   nLin = 3
End If
ImprimeListadoDeComprobacionCabecera = sTexto
End Function
Public Function ImprimeDetalledeCalificacion(pdFecha As Date, _
                pnLinPage As Integer) As String
Dim nLin   As Integer, P As Integer
Dim lnFila As Integer
Dim sTexto As String
Dim sMonedaDes As String
Dim sEstadoDes As String
Dim sCondicDes As String
Dim sProducDes As String
Dim sAgeCod As String
Dim prs As ADODB.Recordset
Dim nSaldoCap As Currency
Dim nIntDeven As Currency
Dim nIntSuspe As Currency
Dim nProvisio As Currency
Dim nSaldoVen As Currency
Dim nCambio As Integer
Dim sProducDescripcion As String
Dim lnCont As Double
Dim lsImpre As String
nLin = 60
Dim dSaldo As DCtaSaldo
Set dSaldo = New DCtaSaldo
Set prs = dSaldo.ObtenerDatosInteresDevSuspenDetalle()
lnFila = 7
ReDim sTotales(3, 8, 4)
Do While Not prs.EOF
      DoEvents
      sTexto = sTexto & ImprimeDetalleDeComprobacionCabecera(pdFecha, pdFecha, 0, P, 0, 1, pdFecha, nLin, pnLinPage)
   
     nCambio = 0
If sProducDes = Mid(prs!cCtaCod, 6, 1) And sCondicDes = prs!cCalGen And sMonedaDes = Mid(prs!cCtaCod, 9, 1) And sEstadoDes = prs!nConsValor And sAgeCod = Mid(prs!cCtaCod, 4, 2) Then
    nCambio = 0
Else
    nCambio = 1
End If
   If lnCont > 0 Then
       If nCambio = 1 Then
            Linea sTexto, CON & BON & Space(40) & "Sub-Total Tipo " & Mid(prs!cConsDescripcion & Space(15), 1, 10) & "  ----->" & Space(9) & PrnVal(nSaldoCap, 16, 2) & Space(48) & PrnVal(nIntDeven, 15, 2) & PrnVal(nIntSuspe, 16, 2) & PrnVal(nProvisio, 16, 2) & PrnVal(nSaldoCap, 13, 2) & BOFF & COFF
            nLin = nLin + 1
            nSaldoCap = 0
            nIntDeven = 0
            nIntSuspe = 0
            nProvisio = 0
            nSaldoVen = 0
        End If
    End If
    If nCambio = 1 Then
        Linea sTexto, prs!cAgeDescripcion
        Linea sTexto, prs!cDescEstado
        Linea sTexto, IIf(Mid(prs!cCtaCod, 9, 1) = "1", "MONEDA NACIONAL", "MONEDA EXTRANEJERA")
        Linea sTexto, prs!cCondicion
        Linea sTexto, prs!cConsDescripcion
        nLin = nLin + 4
    End If
   sMonedaDes = Mid(prs!cCtaCod, 9, 1)
   sEstadoDes = prs!nConsValor
   sCondicDes = prs!cCalGen
   sProducDes = Mid(prs!cCtaCod, 6, 1)
   sProducDescripcion = prs!cConsDescripcion
   sAgeCod = Mid(prs!cCtaCod, 4, 2)
   Linea sTexto, CON & Mid(prs!cCtaCod & Space(18), 1, 18) & Space(2) & Mid(prs!cpersNombre & Space(100), 1, 30) & Space(2) & Mid(prs!cPerscod & Space(20), 1, 18) & Space(1) & PrnVal(prs!nMontoApr, 12, 2) & PrnVal(prs!nSaldoCap, 16, 2) & Space(2) & Format(prs!dFecVig, "YYYY/MM/DD") & Space(2) & Format(prs!dVenc, "YYYY/MM/DD") & Space(2) & IIf(Trim(Format(prs!dPagoCuota, "YYYY/MM/DD")) = "1900/01/01", Space(10), Format(prs!dPagoCuota, "YYYY/MM/DD")) & Space(2) & PrnVal(prs!nDiasAtraso, 4, 0) & Space(1) & PrnVal(prs!nTasaInteres, 5, 2) & PrnVal(prs!nIntDev, 15, 2) & PrnVal(prs!nIntSusp, 16, 2) & PrnVal(prs!nProvision + prs!nProvisionProciclica, 16, 2) & PrnVal(prs!nSaldoCap, 13, 2) & COFF
   
    nSaldoCap = nSaldoCap + prs!nSaldoCap
    nIntDeven = nIntDeven + prs!nIntDev
    nIntSuspe = nIntSuspe + prs!nIntSusp
    nProvisio = nProvisio + prs!nProvision + prs!nProvisionProciclica
    nSaldoVen = nSaldoVen + prs!nCapVencido
      nLin = nLin + 1
      lnCont = lnCont + 1
      If lnCont Mod 600 = 0 Then
         lsImpre = lsImpre & sTexto
         sTexto = ""
      End If
   prs.MoveNext
Loop
     Linea sTexto, CON & BON & Space(27) & "Sub-Total Tipo " & Mid(sProducDescripcion & Space(15), 1, 30) & "  ----->" & Space(12) & PrnVal(nSaldoCap, 16, 2) & Space(48) & PrnVal(nIntDeven, 15, 2) & PrnVal(nIntSuspe, 16, 2) & PrnVal(nProvisio, 16, 2) & PrnVal(nSaldoCap, 13, 2) & BOFF & COFF
        nSaldoCap = 0
        nIntDeven = 0
        nIntSuspe = 0
        nProvisio = 0
        nSaldoVen = 0

ImprimeDetalledeCalificacion = lsImpre & sTexto
End Function

Private Function ImprimeDetalleDeComprobacionCabecera(pdFechaIni As Date, pdFechaFin As Date, pnTpoCambio As Currency, P As Integer, pnTipoBala As Integer, pnMoneda As Integer, pdFecha As Date, ByRef nLin As Integer, ByRef pnLinPage As Integer) As String
Dim sTexto As String
If nLin > pnLinPage - 6 Then
   
   
   Linea sTexto, Cabecera(" LISTADO DE CALIFICACION DE CARTRA -INTERES DEVENGADOS Y EN SUSPENSO - PROVISIONES  ", P, "", 120, , Format(pdFecha, gsFormatoFechaView), lsEmprLogo), 0
   If pnMoneda = 0 Then
      Linea sTexto, Centra(" C O N S O L I D A D O ", 120)
   End If
   
   Linea sTexto, Centra("( DEL " & pdFechaIni & " AL " & pdFechaFin & ")", 120)
     
   Linea sTexto, CON & " ===================================================================================================================================================================================================================="
   Linea sTexto, BON & "  NROCREDITO    ¦ NOMBRE Y/O RAZON SOCIAL        ¦ COD CLIENTE  ¦ MONTOAPRO      ¦ SALDO K       ¦ F.VIGEN   ¦  F.VEN.CRE ¦   PAGADA ¦ MORA T. INT. ¦ DEVENG.     ¦ SUSPENSO     ¦ PROVISION     ¦ % SAL.CUOTA" & BOFF
   Linea sTexto, " ====================================================================================================================================================================================================================" & COFF
   
   nLin = 3
End If
ImprimeDetalleDeComprobacionCabecera = sTexto
End Function

'*** PEAC 20110531
Public Function ImprimeAsientoContResVtaPigno(ByVal psMovNro As String, ByVal pnLinPage As Integer, ByVal pnColPage As Integer, Optional ByVal psTit1 As String, Optional ByVal psCabe1 As String, Optional ByVal psPiePag As String = "9", Optional ByVal psOtraFirma As String = "", Optional psNomCmact As String = "") As String
    On Error GoTo ImprimeAsientoContResVtaPignoErr
    Dim sSQL As String
    Dim rs       As New ADODB.Recordset
    Dim rsDoc       As New ADODB.Recordset
    Dim rsMovs       As New ADODB.Recordset
    Dim nTot     As Currency
    Dim nTotH    As Currency
    Dim sDoc     As String
    Dim sTexto   As String
    Dim sAsiento As String
    Dim nLi      As Integer
    Dim sCabecera As String
    Dim lnCambioFijo As Currency
    Dim lnCambioVenta As Currency
    Dim lnCambioCompra As Currency
    Dim lnCambioVentaSBS As Currency
    Dim lnCambioCompraSBS As Currency
    Dim CON As String
    Dim COFF As String
    Dim lsMovBilletaje As String
    
    Dim lsDesObjeto As String
    Dim lsFecha    As String
    Dim oConect As DConecta
    Dim lsGlosa As String
    Dim lsMovNro As String
    Dim lsOpeCod As String
    Dim sDestino As String
    Dim sOrigen As String
    Dim oDGeneral As DGeneral
    Dim BON As String
    Dim BOFF As String
    Dim oTipoCamb As nTipoCambio
    
    Dim txtMeses As String
    Dim lcFecha As String

    Set oDGeneral = New DGeneral
    Set oTipoCamb = New nTipoCambio
    
    Set oConect = New DConecta
    
    If oConect.AbreConexion = False Then Exit Function
    CON = PrnSet("C+")
    COFF = PrnSet("C-")
    BON = PrnSet("B+")
    BOFF = PrnSet("B-")
    
    sCabecera = ""
    sAsiento = ""
    nLi = 0
    If psMovNro = "" Then
       Exit Function
    End If
       
       psMovNro = Replace(psMovNro, "'", "")
       
       sSQL = "exec stp_sel_BuscaMovNroAsiento '" & psMovNro & "'"
       
       Set rsMovs = oConect.CargaRecordSet(sSQL)
       
       If (rsMovs.EOF And rsMovs.BOF) Then
        Exit Function
       End If
       
       Do While Not rsMovs.EOF
       
        nLi = 0
       
       sSQL = "SELECT * FROM Mov WHERE cMovNro = '" & rsMovs!cMovNro & "'"
       
       Set rs = oConect.CargaRecordSet(sSQL)
       
        lcFecha = Right(psTit1, 10)
        
        txtMeses = Choose(Mid(lcFecha, 4, 2), "Enero", "Febrero", "Marzo", "Abril", _
                                            "Mayo", "Junio", "Julio", "Agosto", _
                                            "Setiembre", "Octubre", "Noviembre", "Diciembre")
       
       lsGlosa = Replace("Joyas Adjudicadas y vendidas en el mes de " & txtMeses & " del " & Right(lcFecha, 4), Chr(13) & oImpresora.gPrnSaltoLinea, " ")
       lsMovNro = rs!cMovNro
       
       lsFecha = GetFechaMov(lsMovNro, True)
       
       lsOpeCod = rs!cOpeCod
       If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Or Left(lsOpeCod, 4) = Left(gOpeME, 4) Or rs!cOpeCod = gContAjusteTipoCambio Then
         lnCambioFijo = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCFijoMes)
         lnCambioVenta = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCVenta)
         lnCambioCompra = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCCompra)
         
         lnCambioVentaSBS = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCPondVenta)
         lnCambioCompraSBS = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCPonderado)
       End If
       
       If rs!nMovFlag = gMovFlagDeExtorno Then
            psTit1 = " A S I E N T O  C O N T A B L E   D E   E X T O R N O"
       ElseIf rs!nMovFlag = gMovFlagEliminado Then
            psTit1 = " **** ASIENTO ELIMINADO NO VALIDO PARA CONTABILIDAD **** "
       End If
       'cabecera
       If Left(lsOpeCod, 3) = "622" Or Left(lsOpeCod, 3) = "632" Then
            sAsiento = sAsiento & ImpreCabAsiento(pnColPage, CDate(lsFecha), gsNomCmac, lsOpeCod, lsMovNro, psTit1, False, , , nLi) & oImpresora.gPrnSaltoLinea
       Else
            sAsiento = sAsiento & ImpreCabAsiento(pnColPage, CDate(lsFecha), gsNomCmac, lsOpeCod, lsMovNro, psTit1, True, , , nLi) & oImpresora.gPrnSaltoLinea
       End If
       If psCabe1 <> "" Then
          sAsiento = sAsiento & psCabe1 & oImpresora.gPrnSaltoLinea
          nLi = nLi + 4
       Else
           'BUSCA LAS PERSONA INVOLUCRADAS EN EL MOVIMIENTO
            sSQL = " SELECT  DISTINCT M.CMOVNRO,  " _
                    & "         CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                    & "                 WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                    & "         P.cPersCod, " _
                    & "         Convert(Char(40),P.CPERSNOMBRE) as cPersNombre, " _
                    & "         MOI.CPersCod " _
                    & " FROM    MOV M JOIN MOVCTA MC    ON M.NMOVNRO = MC.NMOVNRO " _
                    & "         JOIN MOVGasto MOI         ON MOI.NMOVNRO = M.NMOVNRO " _
                    & "         JOIN PERSONA P          ON P.CPERSCOD = MOI.CPERSCOD " _
                    & " WHERE   M.cMovNro = '" & lsMovNro & "' "
              
            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = ""
            Do While Not rs.EOF
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & PstaNombre(rs!cpersNombre)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & PstaNombre(rs!cpersNombre)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat("Persona(s) Origen  ", 25) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat("Persona(s) Destino ", 25) & " :" & sDestino & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
            sSQL = " SELECT  DISTINCT M.CMOVNRO, " _
                & "         CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                 WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "         P.cPersCod , CI.cCtaIFCod , " _
                & "         Convert(Char(40),CTPOI.cConsDescripcion ) as Entidad, " _
                & "         Convert(Char(40),P.CPERSNOMBRE) as cPersNombre, " _
                & "         Convert(char(50),C.cConsDescripcion + ' ' + CI.cCtaIFDesc) as DescCtaIF, " _
                & "         MO.COBJETOCOD " _
                & " FROM    MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "         JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "         JOIN MOVOBJIF MOI       ON MOI.NMOVNRO = MO.NMOVNRO AND MOI.NMOVITEM = MO.NMOVITEM " _
                & "         JOIN CTAIF CI           ON CI.CCTAIFCOD = MOI.CCTAIFCOD AND MOI.CPERSCOD = CI.CPERSCOD AND MOI.CIFTPO = CI.CIFTPO  " _
                & "         JOIN CONSTANTE C        ON C.nConsValor = SUBSTRING(CI.CCTAIFCOD,2,1) AND C.nCONSCOD LIKE '" & gCGTipoCtaIF & "' " _
                & "         JOIN INSTITUCIONFINANC I    ON I.CPERSCOD = CI.CPERSCOD AND I.CIFTPO = CI.CIFTPO " _
                & "         JOIN PERSONA P          ON P.CPERSCOD = I.CPERSCOD " _
                & "         JOIN CONSTANTE CTPOI        ON CTPOI.nConsValor = I.cIFTpo AND CTPOI.nCONSCOD LIKE '" & gCGTipoIF & "' " _
                & " WHERE M.cMovNro = '" & lsMovNro & "' AND MO.cObjetoCod ='" & Format(ObjEntidadesFinancieras, "00") & "'"
            
            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cpersNombre) & " " & Trim(rs!DescCtaIF)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cpersNombre) & " " & Trim(rs!DescCtaIF)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat("Cuenta(s) Origen", 20) & " :" + PrnSet("C+") + sOrigen + PrnSet("C-") + oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat("Cuenta(s) Destino", 20) & " :" + PrnSet("C+") + sDestino + PrnSet("C-") + oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
            sSQL = "    SELECT  DISTINCT M.CMOVNRO, " _
                & "             CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                     WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "             MOA.cAreaCod + MOA.CAGECOD AS cAreaCod, A.cAreaDescripcion + ' ' + ISNULL(AG.CAGEDESCRIPCION,'' )  as cAreaDescripcion , MO.COBJETOCOD, O.cObjetoDesc " _
                & "     FROM    MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "             JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "             JOIN MovObjAreaAgencia MOA         ON MOA.NMOVNRO = MO.NMOVNRO AND MOA.NMOVITEM = MO.NMOVITEM " _
                & "             JOIN AREAS A            ON A.cAreaCod = MOA.cAreaCod LEFT JOIN AGENCIAS AG ON AG.CAGECOD = MOA.CAGECOD " _
                & "             JOIN OBJETO O ON O.cObjetoCod  = MO.cObjetoCod " _
                & "     WHERE   M.cMovNro = '" & lsMovNro & "' AND Mo.cObjetoCod IN ('" & Format(ObjCMACAgenciaArea, "00") & "','" & Format(ObjCMACArea, "00") & "')"

            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                lsDesObjeto = Trim(rs!cObjetoDesc)
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!cAreaDescripcion)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!cAreaDescripcion)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Origen ", 20) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Destino ", 20) & " :" & sDestino & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If

            sSQL = " SELECT    DISTINCT M.CMOVNRO, " _
                & "             CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                     WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "             MO.COBJETOCOD, O.cObjetoDesc " _
                & "   FROM      MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "             JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "             JOIN OBJETO O ON O.cObjetoCod  = MO.cObjetoCod " _
                & "   WHERE     M.cMovNro='" & lsMovNro & "' " _
                & "             AND MO.cObjetoCod NOT IN ('" & Format(ObjCMACAgencias, "00") & "','" & Format(ObjCMACAgenciaArea, "00") & "','" & Format(ObjDescomEfectivo, "00") & "','" & Format(ObjEntidadesFinancieras, "00") & "','" & Format(ObjBienesServicios, "00") & "','" & Format(ObjPersona, "00") & "')"
            
            Set rs = oConect.CargaRecordSet(sSQL)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                lsDesObjeto = "OBJETOS"
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cObjetoCod) & "-" & Trim(rs!cObjetoDesc)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cObjetoCod) & "-" & Trim(rs!cObjetoDesc)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Origen ", 20) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Destino", 20) & " :" & sDestino & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
            
        End If ' del pscab1

       'Tipo de Cambio
        sSQL = "SELECT nMovTpoCambio, m.cOpeCod FROM MovTpoCambio MC JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE M.cMovNro = '" & lsMovNro & "'"
        Set rs = oConect.CargaRecordSet(sSQL)
        If Not rs.EOF Then
            If rs!cOpeCod = gContAjusteTipoCambio Then
            Else
               sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "       T.C.C SBS : " & Format(lnCambioCompraSBS, "##,###,##0.000") & "     T.C.V SBS : " & Format(lnCambioVentaSBS, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
               'sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "      T.C.Mercado: " & Format(rs!nMovTpoCambio, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
            End If
            nLi = nLi + 1
        Else
            If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Then
                sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "     T.C.Compra : " & Format(lnCambioCompra, "##,###,##0.000") & "     T.C.Venta : " & Format(lnCambioVenta, "##,###,##0.000") & "     T.C.C SBS : " & Format(lnCambioCompraSBS, "##,###,##0.000") & "     T.C.V SBS : " & Format(lnCambioVentaSBS, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
            End If
        End If
       nLi = nLi + 1
       'Datos de Documentos
       sSQL = "SELECT b.cDocAbrev, a.cDocNro, a.dDocFecha FROM MovDoc a join mov m on m.nmovnro =a.nmovnro  , " & vsBaseComunes & "Documento b " _
            & "WHERE m.cMovNro = '" & lsMovNro & "' and b.nDocTpo = a.nDocTpo "

        Set rsDoc = oConect.CargaRecordSet(sSQL)
        If Not rsDoc.EOF Then
          sDoc = ""
          nLi = 1
          Do While Not rsDoc.EOF
             sDoc = sDoc & rsDoc!cDocAbrev & "_" & rsDoc!cDocNro & " " & rsDoc!dDocFecha & "  "
             If nLi Mod 4 = 0 Then
                sDoc = sDoc + oImpresora.gPrnSaltoLinea
                nLi = nLi + 1
             End If
             rsDoc.MoveNext
          Loop
          sAsiento = sAsiento + ImpreFormat("Documentos", 20) & " :" + PrnSet("C+") + sDoc + PrnSet("C-") + oImpresora.gPrnSaltoLinea
          nLi = nLi + 1
        End If
        rsDoc.Close
       'Glosa
       sAsiento = sAsiento & ImpreGlosa(lsGlosa, pnColPage, ImpreFormat("Glosa       ", 20) & " :", , , nLi) + oImpresora.gPrnSaltoLinea
       nLi = nLi + 1
       lsMovBilletaje = ""
       lsMovBilletaje = ImprimeBilletaje(Mid(lsOpeCod, 3, 1), lsMovNro, pnColPage, , nLi)
       If lsMovBilletaje <> "" Then
            sAsiento = sAsiento + lsMovBilletaje
       End If
    'Cabecera 2 de Asiento
    sCabecera = ""
    sCabecera = sCabecera + CON
    sCabecera = sCabecera + "====================================================================================================================================" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "Itm C U E N T A   C O N T A B L E                                                                           DEBE            HABER" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "    Codigo              Descripcion" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "------------------------------------------------------------------------------------------------------------------------------------" & COFF & oImpresora.gPrnSaltoLinea
    'MovCta para MN
    sSQL = "SELECT mc.nMovItem, mc.cCtaContCod, ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),'') cCtaContDesc, mc.nMovImporte " _
         & "FROM   MovCta mc JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE  M.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
    Set rs = oConect.CargaRecordSet(sSQL)
    If rs.EOF Then
       'MsgBox "No existen datos de Movimiento ", vbInformation, "Error"
       Exit Function
    End If
    nTot = 0
    nTotH = 0
    sAsiento = sAsiento + sCabecera
    nLi = nLi + 4
    sAsiento = sAsiento + BON & " EN MONEDA NACIONAL " + oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento + "--------------------" & BOFF + oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento & oImpresora.gPrnSaltoLinea
    Do While Not rs.EOF
       DoEvents
       If nLi > pnLinPage - 6 Then
           nLi = 5
           sAsiento = sAsiento + oImpresora.gPrnSaltoPagina & sCabecera
       End If
       sAsiento = sAsiento + CON & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " & Justifica(rs!cCtaContDesc, 71) & " " _
            & Right(Space(16) & IIf(rs!nMovImporte > 0, Format(rs!nMovImporte, "##,###,##0.00"), ""), 16) & " " _
            & Right(Space(16) & IIf(rs!nMovImporte < 0, Format(rs!nMovImporte * -1, "##,###,##0.00"), ""), 16) _
            & COFF & oImpresora.gPrnSaltoLinea
       If rs!nMovImporte > 0 Then
          nTot = nTot + Val(rs!nMovImporte)
       Else
          nTotH = nTotH + Val(rs!nMovImporte) * -1
       End If
       nLi = nLi + 1
       rs.MoveNext
    Loop
    sAsiento = sAsiento + CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF & oImpresora.gPrnSaltoLinea
    nLi = nLi + 2

    sSQL = "SELECT mc.nMovItem, mc.cCtaContCod, LTRIM(RTRIM(ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),''))) cCtaContDesc, me.nmovMEimporte " _
         & "FROM   mov m join MovCta mc on mc.nmovnro=m.nmovnro JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
         & "WHERE  m.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
    Set rs = oConect.CargaRecordSet(sSQL)
    sTexto = ""
    nTot = 0
    nTotH = 0
    If Not rs.EOF Then
       sAsiento = sAsiento + BON & " EN MONEDA EXTRANJERA " & oImpresora.gPrnSaltoLinea
       sAsiento = sAsiento + "----------------------" & BOFF & oImpresora.gPrnSaltoLinea
       nLi = nLi + 2
    End If
    Do While Not rs.EOF
       DoEvents
       If nLi > pnLinPage - 6 Then
          nLi = 5
          sTexto = sTexto + oImpresora.gPrnSaltoPagina
          sTexto = sTexto + sCabecera
       End If
       sTexto = sTexto + CON & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " _
              & Justifica(rs!cCtaContDesc, 71) & " " _
              & IIf(rs!nMovMEImporte > 0, PrnVal(rs!nMovMEImporte, 16, 2), Space(16)) & " " _
              & IIf(rs!nMovMEImporte < 0, PrnVal(rs!nMovMEImporte * -1, 16, 2), Space(16)) _
              & COFF & oImpresora.gPrnSaltoLinea
       If rs!nMovMEImporte > 0 Then
          nTot = nTot + Val(rs!nMovMEImporte)
       Else
          nTotH = nTotH + Val(rs!nMovMEImporte) * -1
       End If
       nLi = nLi + 1
       rs.MoveNext
    Loop
    If nTot > 0 Or nTotH > 0 Then
       sAsiento = sAsiento & sTexto
       sAsiento = sAsiento + CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
       sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF + oImpresora.gPrnSaltoLinea
       nLi = nLi + 3
    End If
    
    If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
    
    If lsOpeCod <> 421113 Then
        sAsiento = sAsiento + CON & "====================================================================================================================================" & COFF & oImpresora.gPrnSaltoLinea
        
        sAsiento = sAsiento + ImprePiePag(pnColPage, psPiePag, psOtraFirma)
    End If
    sAsiento = sAsiento + "   " + oImpresora.gPrnSaltoLinea
    
    sAsiento = sAsiento + "   " + oImpresora.gPrnSaltoPagina
    
    rsMovs.MoveNext
    Loop

    Set oDGeneral = Nothing
    Set oConect = Nothing
    
    'End If
    
    ImprimeAsientoContResVtaPigno = sAsiento
    
    Exit Function
ImprimeAsientoContResVtaPignoErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeAsientoContResVtaPigno Method")
End Function

'JACA 20111017****************************************************************
Public Sub GeneraExcelAnexo12_II(prs As Recordset, pdFecha As Date, psAppPath As String)
    
        
        Dim sPathAnexo12_II As String
        Dim sFormatoAnexo12_II As String
        Dim fs As New Scripting.FileSystemObject
        Dim obj_Excel As Object, Libro As Object, Hoja As Object
        
        On Error GoTo error_sub
          
       sPathAnexo12_II = psAppPath & "\Spooler\ANEXO12_II_" + UCase(Left(Format(pdFecha, "MMMM"), 3)) + CStr(Year(pdFecha)) + ".xls"

        
        If fs.FileExists(sPathAnexo12_II) Then
'            If ArchivoEstaAbierto(sPathAnexo12_II) Then
'                If MsgBox("Debe Cerrar el Archivo:" + fs.GetFileName(sPathAnexo12_II) + " para continuar", vbRetryCancel) = vbCancel Then
'                   Exit Sub
'                End If
'            End If
            Do While ArchivoEstaAbierto(sPathAnexo12_II)
                If MsgBox("Debe Cerrar el Archivo:" + fs.GetFileName(sPathAnexo12_II) + " para continuar", vbRetryCancel) = vbCancel Then
                   Exit Sub
                End If
            Loop
            fs.DeleteFile sPathAnexo12_II, True
        End If
              
        sFormatoAnexo12_II = psAppPath + "\FormatoCarta\ANEXO_12_II.xls"
        
        If Len(Dir(sFormatoAnexo12_II)) = 0 Then
           MsgBox "No se Pudo Encontrar el Archivo:" & sFormatoAnexo12_II, vbCritical
           Exit Sub
        End If
        
        Set obj_Excel = CreateObject("Excel.Application")
        obj_Excel.DisplayAlerts = False
        Set Libro = obj_Excel.Workbooks.Open(sFormatoAnexo12_II)
        Set Hoja = Libro.ActiveSheet
        
        Dim celda As Excel.Range
        Dim oCtaCont As DbalanceCont
        Dim rsCtaCont As ADODB.Recordset
        
        Set oCtaCont = New DbalanceCont
        Set rsCtaCont = New ADODB.Recordset
       
       ' Fecha del ANEXO
        Set celda = obj_Excel.Range("C8")
        celda.value = "AL " + Format(Day(pdFecha), "00") + " DE " + UCase(Format(pdFecha, "MMMM")) + " DEL " + CStr(Year(pdFecha))
        
        '****************************MONEDA NACIONAL****************************************
        If prs!nMoneda = 1 Then 'verifica si la fila es moneda nacional
            Set celda = obj_Excel.Range("D27")
            celda.value = prs!Fec_Emision
            Set celda = obj_Excel.Range("E27")
            celda.value = prs!Plazo
            Set celda = obj_Excel.Range("G27")
            celda.value = prs!Monto
            Set celda = obj_Excel.Range("J27")
            celda.value = prs!Porcion
            
            If prs.RecordCount > 1 Then
                prs.MoveNext
            End If
        End If
        '****************************MONEDA ESTRANJERA****************************************
        
        If prs!nMoneda = 2 Then
            Set celda = obj_Excel.Range("D32")
            celda.value = prs!Fec_Emision
            Set celda = obj_Excel.Range("E32")
            celda.value = prs!Plazo
            Set celda = obj_Excel.Range("G32")
            celda.value = prs!Monto
            Set celda = obj_Excel.Range("J32")
            celda.value = prs!Porcion
        End If
        

        'guarda el archivo
        Hoja.SaveAs sPathAnexo12_II
        'obj_Excel.Visible = True
        Libro.Close
        obj_Excel.Quit
        
        Set Hoja = Nothing
        Set Libro = Nothing
        Set obj_Excel = Nothing
        
        'abre y muestra el archivo
        
        Dim m_Excel As New Excel.Application
        m_Excel.Workbooks.Open (sPathAnexo12_II)
        m_Excel.Visible = True
        
Exit Sub
error_sub:
        MsgBox TextErr(Err.Description), vbInformation, "Aviso"
        Set Libro = Nothing
        Set obj_Excel = Nothing
        Set Hoja = Nothing
        
End Sub
Private Function ArchivoEstaAbierto(ByVal Ruta As String) As Boolean
On Error GoTo HayErrores
Dim f As Integer
   f = FreeFile
   Open Ruta For Append As f
   Close f
   ArchivoEstaAbierto = False
   Exit Function
HayErrores:
   If Err.Number = 70 Then
      ArchivoEstaAbierto = True
   Else
      Err.Raise Err.Number
   End If
End Function
'JACA END*********************************************************************

'** Juez 20120712 ************************************************************' Adecuado solo para Operaciones de Contingencias, si se reutiliza adecuarlo para otras operaciones
Public Function VistaPreviaAsientoContable(ByVal psMovNro As String, ByVal pnLinPage As Integer, ByVal pnColPage As Integer, Optional ByVal psTit1 As String, Optional ByVal psCabe1 As String, Optional ByVal psPiePag As String = "9", Optional ByVal psOtraFirma As String = "", Optional psNomCmact As String = "", Optional pbHistoCtaCont As Boolean = False, Optional ByVal lsOpeCod As String, Optional ByVal psNroRegistro As String, Optional ByVal pnMoneda As Integer, Optional ByVal DCtaCont As String = "", Optional ByVal HCtaCont As String = "", Optional ByVal nMovImporte As Currency = "0.00", Optional ByVal DCtaContDem As String = "", Optional ByVal HCtaContDem As String = "", Optional ByVal nMovImporteDem As Currency = "0.00") As String
    On Error GoTo VistaPreviaAsientoContableErr
    Dim sSQL As String
    Dim rs       As New ADODB.Recordset
    Dim rsDoc       As New ADODB.Recordset
    Dim nTot     As Currency
    Dim nTotH    As Currency
    Dim sDoc     As String
    Dim sTexto   As String
    Dim sAsiento As String
    Dim nLi      As Integer
    Dim sCabecera As String
    Dim lnCambioFijo As Currency
    Dim lnCambioVenta As Currency
    Dim lnCambioCompra As Currency
    Dim lnCambioVentaSBS As Currency
    Dim lnCambioCompraSBS As Currency
    Dim CON As String
    Dim COFF As String
    Dim lsMovBilletaje As String
    
    Dim lsDesObjeto As String
    Dim lsFecha    As String
    Dim oConect As DConecta
    Dim lsGlosa As String
    Dim lsMovNro As String
    'Dim lsOpeCod As String
    Dim sDestino As String
    Dim sOrigen As String
    Dim oDGeneral As DGeneral
    Dim BON As String
    Dim BOFF As String
    Dim oTipoCamb As nTipoCambio
    
    Set oDGeneral = New DGeneral
    Set oTipoCamb = New nTipoCambio
    
    Set oConect = New DConecta
    
    If oConect.AbreConexion = False Then Exit Function
    CON = PrnSet("C+")
    COFF = PrnSet("C-")
    BON = PrnSet("B+")
    BOFF = PrnSet("B-")
    
    sCabecera = ""
    sAsiento = ""
    nLi = 0
    
    'PARA PARAMETROS, psMovNro,cOpeCod
    
    'If psMovNro = "" Then
    '   Exit Function
    'End If
       lsFecha = GetFechaMov(psMovNro, True)
       'sSql = "SELECT * FROM Mov WHERE cMovNro = '" & psMovNro & "'"
       'Set rs = oConect.CargaRecordSet(sSql)
       'If rs.EOF Then
       '   Exit Function
       'End If
       'lsGlosa = Replace(rs!cMovdesc, Chr(13) & oImpresora.gPrnSaltoLinea, " ")
       'lsMovNro = psMovNro
       'lsOpeCod = rs!cOpeCod
       'If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Or Left(lsOpeCod, 4) = Left(gOpeME, 4) Or rs!cOpeCod = gContAjusteTipoCambio Then
         lnCambioFijo = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCFijoMes)
       '  lnCambioVenta = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCVenta)
       '  lnCambioCompra = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCCompra)
       '
       '  lnCambioVentaSBS = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCPondVenta)
       '  lnCambioCompraSBS = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCPonderado)
       'End If
       psTit1 = " V I S T A  P R E V I A  A S I E N T O  C O N T A B L E "
       'If rs!nMovFlag = gMovFlagDeExtorno Then
       '     psTit1 = " A S I E N T O  C O N T A B L E   D E   E X T O R N O"
       'ElseIf rs!nMovFlag = gMovFlagEliminado Then
       '     psTit1 = " **** ASIENTO ELIMINADO NO VALIDO PARA CONTABILIDAD **** "
       'End If
       '
       'If Left(lsOpeCod, 3) = "622" Or Left(lsOpeCod, 3) = "632" Then
       '     sAsiento = sAsiento & ImpreCabAsiento(pnColPage, CDate(lsFecha), gsNomCmac, lsOpeCod, lsMovNro, psTit1, False, , , nLi) & oImpresora.gPrnSaltoLinea
       'Else
            sAsiento = sAsiento & ImpreCabAsiento(pnColPage, CDate(lsFecha), gsNomCmac, lsOpeCod, lsMovNro, psTit1, True, , , nLi) & oImpresora.gPrnSaltoLinea
       'End If
       'If psCabe1 <> "" Then
       '   sAsiento = sAsiento & psCabe1 & oImpresora.gPrnSaltoLinea
       '   nLi = nLi + 4
       'Else
           'BUSCA LAS PERSONA INVOLUCRADAS EN EL MOVIMIENTO
            'sSql = " SELECT  DISTINCT M.CMOVNRO,  " _
            '        & "         CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
            '        & "                 WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
            '        & "         P.cPersCod, " _
            '        & "         Convert(Char(40),P.CPERSNOMBRE) as cPersNombre, " _
            '        & "         MOI.CPersCod " _
            '        & " FROM    MOVTMP M JOIN MOVCTA MC    ON M.NMOVNRO = MC.NMOVNRO " _
            '        & "         JOIN MOVGastoTmp MOI         ON MOI.NMOVNRO = M.NMOVNRO " _
            '        & "         JOIN PERSONA P          ON P.CPERSCOD = MOI.CPERSCOD " _
            '        & " WHERE   M.cMovNro = '" & lsMovNro & "' "
              
            'Set rs = oConect.CargaRecordSet(sSql)
            'sDestino = "": sOrigen = ""
            'Do While Not rs.EOF
            '    If rs!cDH = "D" Then
            '        sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & PstaNombre(rs!cpersNombre)
            '    Else
            '        sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & PstaNombre(rs!cpersNombre)
            '    End If
            '    nLi = nLi + 1
            '    rs.MoveNext
            'Loop
            'If sOrigen <> "" Then
            '    sAsiento = sAsiento & ImpreFormat("Persona(s) Origen  ", 25) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
            '    nLi = nLi + 1
            'End If
            'If sDestino <> "" Then
            '    sAsiento = sAsiento & ImpreFormat("Persona(s) Destino ", 25) & " :" & sDestino & oImpresora.gPrnSaltoLinea
            '    nLi = nLi + 1
            'End If
            
            'sSql = " SELECT  DISTINCT M.CMOVNRO, " _
            '    & "         CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
            '    & "                 WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
            '    & "         P.cPersCod , CI.cCtaIFCod , " _
            '    & "         Convert(Char(40),CTPOI.cConsDescripcion ) as Entidad, " _
            '    & "         Convert(Char(40),P.CPERSNOMBRE) as cPersNombre, " _
            '    & "         Convert(char(50),C.cConsDescripcion + ' ' + CI.cCtaIFDesc) as DescCtaIF, " _
            '    & "         MO.COBJETOCOD " _
            '    & " FROM    MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
            '    & "         JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
            '    & "         JOIN MOVOBJIF MOI       ON MOI.NMOVNRO = MO.NMOVNRO AND MOI.NMOVITEM = MO.NMOVITEM " _
            '    & "         JOIN CTAIF CI           ON CI.CCTAIFCOD = MOI.CCTAIFCOD AND MOI.CPERSCOD = CI.CPERSCOD AND MOI.CIFTPO = CI.CIFTPO  " _
            '    & "         JOIN CONSTANTE C        ON C.nConsValor = SUBSTRING(CI.CCTAIFCOD,2,1) AND C.nCONSCOD LIKE '" & gCGTipoCtaIF & "' " _
            '    & "         JOIN INSTITUCIONFINANC I    ON I.CPERSCOD = CI.CPERSCOD AND I.CIFTPO = CI.CIFTPO " _
            '    & "         JOIN PERSONA P          ON P.CPERSCOD = I.CPERSCOD " _
            '    & "         JOIN CONSTANTE CTPOI        ON CTPOI.nConsValor = I.cIFTpo AND CTPOI.nCONSCOD LIKE '" & gCGTipoIF & "' " _
            '    & " WHERE M.cMovNro = '" & lsMovNro & "' AND MO.cObjetoCod ='" & Format(ObjEntidadesFinancieras, "00") & "'"
            
            'Set rs = oConect.CargaRecordSet(sSql)
            'sDestino = "": sOrigen = "": lsDesObjeto = ""
            'Do While Not rs.EOF
            '    If rs!cDH = "D" Then
            '        sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cpersNombre) & " " & Trim(rs!DescCtaIF)
            '    Else
            '        sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cpersNombre) & " " & Trim(rs!DescCtaIF)
            '    End If
            '    nLi = nLi + 1
            '    rs.MoveNext
            'Loop
            'If sOrigen <> "" Then
            '    sAsiento = sAsiento & ImpreFormat("Cuenta(s) Origen", 20) & " :" + PrnSet("C+") + sOrigen + PrnSet("C-") + oImpresora.gPrnSaltoLinea
            '    nLi = nLi + 1
            'End If
            'If sDestino <> "" Then
            '    sAsiento = sAsiento & ImpreFormat("Cuenta(s) Destino", 20) & " :" + PrnSet("C+") + sDestino + PrnSet("C-") + oImpresora.gPrnSaltoLinea
            '    nLi = nLi + 1
            'End If
            
            'sSql = "    SELECT  DISTINCT M.CMOVNRO, " _
            '    & "             CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
            '    & "                     WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
            '    & "             MOA.cAreaCod + MOA.CAGECOD AS cAreaCod, A.cAreaDescripcion + ' ' + ISNULL(AG.CAGEDESCRIPCION,'' )  as cAreaDescripcion , MO.COBJETOCOD, O.cObjetoDesc " _
            '    & "     FROM    MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
            '    & "             JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
            '    & "             JOIN MovObjAreaAgencia MOA         ON MOA.NMOVNRO = MO.NMOVNRO AND MOA.NMOVITEM = MO.NMOVITEM " _
            '    & "             JOIN AREAS A            ON A.cAreaCod = MOA.cAreaCod LEFT JOIN AGENCIAS AG ON AG.CAGECOD = MOA.CAGECOD " _
            '    & "             JOIN OBJETO O ON O.cObjetoCod  = MO.cObjetoCod " _
            '    & "     WHERE   M.cMovNro = '" & lsMovNro & "' AND Mo.cObjetoCod IN ('" & Format(ObjCMACAgenciaArea, "00") & "','" & Format(ObjCMACArea, "00") & "')"

            'Set rs = oConect.CargaRecordSet(sSql)
            'sDestino = "": sOrigen = "": lsDesObjeto = ""
            'Do While Not rs.EOF
            '    lsDesObjeto = Trim(rs!cObjetoDesc)
            '    If rs!cDH = "D" Then
            '        sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!cAreaDescripcion)
            '    Else
            '        sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!cAreaDescripcion)
            '    End If
            '    nLi = nLi + 1
            '    rs.MoveNext
            'Loop
            'If sOrigen <> "" Then
            '    sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Origen ", 20) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
            '    nLi = nLi + 1
            'End If
            'If sDestino <> "" Then
            '    sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Destino ", 20) & " :" & sDestino & oImpresora.gPrnSaltoLinea
            '    nLi = nLi + 1
            'End If

            'sSql = " SELECT    DISTINCT M.CMOVNRO, " _
            '    & "             CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
            '    & "                     WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
            '    & "             MO.COBJETOCOD, O.cObjetoDesc " _
            '    & "   FROM      MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
            '    & "             JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
            '    & "             JOIN OBJETO O ON O.cObjetoCod  = MO.cObjetoCod " _
            '    & "   WHERE     M.cMovNro='" & lsMovNro & "' " _
            '    & "             AND MO.cObjetoCod NOT IN ('" & Format(ObjCMACAgencias, "00") & "','" & Format(ObjCMACAgenciaArea, "00") & "','" & Format(ObjDescomEfectivo, "00") & "','" & Format(ObjEntidadesFinancieras, "00") & "','" & Format(ObjBienesServicios, "00") & "','" & Format(ObjPersona, "00") & "')"
            
            'Set rs = oConect.CargaRecordSet(sSql)
            'sDestino = "": sOrigen = "": lsDesObjeto = ""
            'Do While Not rs.EOF
            '    lsDesObjeto = "OBJETOS"
            '    If rs!cDH = "D" Then
            '        sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cObjetoCod) & "-" & Trim(rs!cObjetoDesc)
            '    Else
            '        sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(rs!cObjetoCod) & "-" & Trim(rs!cObjetoDesc)
            '    End If
            '    nLi = nLi + 1
            '    rs.MoveNext
            'Loop
            'If sOrigen <> "" Then
            '    sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Origen ", 20) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
            '    nLi = nLi + 1
            'End If
            'If sDestino <> "" Then
            '    sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Destino", 20) & " :" & sDestino & oImpresora.gPrnSaltoLinea
            '    nLi = nLi + 1
            'End If
            
        'End If ' del pscab1

       'If lsOpeCod <> "500520" And lsOpeCod <> "500510" Then
       ' 'Tipo de Cambio
       '  sSql = "SELECT nMovTpoCambio, m.cOpeCod FROM MovTpoCambio MC JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE M.cMovNro = '" & lsMovNro & "'"
       '  Set rs = oConect.CargaRecordSet(sSql)
       '  If Not rs.EOF Then
       '      If rs!cOpeCod = gContAjusteTipoCambio Then
       '      Else
       '         sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "       T.C.C SBS : " & Format(lnCambioCompraSBS, "##,###,##0.000") & "     T.C.V SBS : " & Format(lnCambioVentaSBS, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
       '         'sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "      T.C.Mercado: " & Format(rs!nMovTpoCambio, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
       '      End If
       '      nLi = nLi + 1
       '  Else
       '      If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Then
       '          sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "     T.C.Compra : " & Format(lnCambioCompra, "##,###,##0.000") & "     T.C.Venta : " & Format(lnCambioVenta, "##,###,##0.000") & "     T.C.C SBS : " & Format(lnCambioCompraSBS, "##,###,##0.000") & "     T.C.V SBS : " & Format(lnCambioVentaSBS, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
       '          nLi = nLi + 1
       '      End If
       '  End If
       'End If

        If lsOpeCod = "500520" Then
            sAsiento = sAsiento & "  Fecha: " & lsFecha & oImpresora.gPrnSaltoLinea '& Space(132 - 45) & "Usuario: " & FillText(gsCodUser, 5, " ") & oImpresora.gPrnSaltoLinea
            'sAsiento = sAsiento & Space(132 - 50) & "Area: " & FillText(gsCodUser, 30, " ") & oImpresora.gPrnSaltoLinea
        End If
        If lsOpeCod = "500510" Then
            Dim oConting As DContingencia
            Dim rsConting As ADODB.Recordset
            Set oConting = New DContingencia
            Set rsConting = oConting.ObtieneUltimoInformeTecConting(psNroRegistro)
            sAsiento = sAsiento & "  Informe N*: " & oImpresora.gPrnSaltoLinea '& rsConting!cNroInforme & oImpresora.gPrnSaltoLinea
            sAsiento = sAsiento & "  Fecha de Informe: " & oImpresora.gPrnSaltoLinea '& Format(rsConting!dFechaRegInfTec, "dd/mm/yyyy") & oImpresora.gPrnSaltoLinea
        End If

        nLi = nLi + 1
       'Datos de Documentos
       'sSql = "SELECT b.cDocAbrev, a.cDocNro, a.dDocFecha FROM MovDoc a join mov m on m.nmovnro =a.nmovnro  , " & vsBaseComunes & "Documento b " _
       '     & "WHERE m.cMovNro = '" & lsMovNro & "' and b.nDocTpo = a.nDocTpo "

        'Set rsDoc = oConect.CargaRecordSet(sSql)
        'If Not rsDoc.EOF Then
        '  sDoc = ""
        '  nLi = 1
        '  Do While Not rsDoc.EOF
        '     sDoc = sDoc & rsDoc!cDocAbrev & "_" & rsDoc!cDocNro & " " & rsDoc!dDocFecha & "  "
        '     If nLi Mod 4 = 0 Then
        '        sDoc = sDoc + oImpresora.gPrnSaltoLinea
        '        nLi = nLi + 1
        '     End If
        '     rsDoc.MoveNext
        '  Loop
        '  sAsiento = sAsiento + ImpreFormat("Documentos", 20) & " :" + PrnSet("C+") + sDoc + PrnSet("C-") + oImpresora.gPrnSaltoLinea
        '  nLi = nLi + 1
        'End If
        'rsDoc.Close
       'Glosa
        If lsOpeCod <> "500520" And lsOpeCod <> "500510" Then
            sAsiento = sAsiento & ImpreGlosa(lsGlosa, pnColPage, ImpreFormat("Glosa       ", 20) & " :", , , nLi) + oImpresora.gPrnSaltoLinea
        Else
            sAsiento = sAsiento & "  Glosa: " & lsGlosa & oImpresora.gPrnSaltoLinea
        End If
       nLi = nLi + 1
       'lsMovBilletaje = ""
       'lsMovBilletaje = ImprimeBilletaje(Mid(lsOpeCod, 3, 1), lsMovNro, pnColPage, , nLi)
       'If lsMovBilletaje <> "" Then
       '     sAsiento = sAsiento + lsMovBilletaje
       'End If
    'Cabecera 2 de Asiento
    sCabecera = sCabecera + CON
    sCabecera = sCabecera + "====================================================================================================================================" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "Itm C U E N T A   C O N T A B L E                                                                           DEBE            HABER" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "    Codigo              Descripcion" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "------------------------------------------------------------------------------------------------------------------------------------" & COFF & oImpresora.gPrnSaltoLinea
    'MovCta para MN
    
    'JACA 20111227 Modificado para obtener cCtaContDesc Historica**************************
    'sSql = "SELECT mc.nMovItem, mc.cCtaContCod," + IIf(pbHistoCtaCont = False, "ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),'')", "ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),'')") + " cCtaContDesc, mc.nMovImporte " _
    '    & "FROM   MovCta mc JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE  M.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
'    sSQL = "SELECT mc.nMovItem, mc.cCtaContCod," + IIf(pbHistoCtaCont = False, "ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),'')", "ISNULL(dbo.Histo_GetCtaContDesc(mc.cCtaContCod,2,1),'')") + " cCtaContDesc, mc.nMovImporte " _
'         & "FROM   MovCta mc JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE  M.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
    'JACA END******************************************************************
    
    If DCtaCont <> "" And HCtaCont <> "" Then
        sSQL = "SELECT 1 nMovItem, '" & DCtaCont & "' cCtaContCod, ISNULL(dbo.GetCtaContDesc('" & DCtaCont & "',2,1),'') cCtaContDesc, " & nMovImporte * IIf(pnMoneda = gMonedaExtranjera, lnCambioFijo, 1) & " nMovImporte"
        sSQL = sSQL + " UNION "
        sSQL = sSQL + "SELECT 2 nMovItem, '" & HCtaCont & "' cCtaContCod, ISNULL(dbo.GetCtaContDesc('" & HCtaCont & "',2,1),'') cCtaContDesc, " & nMovImporte * IIf(pnMoneda = gMonedaExtranjera, lnCambioFijo, 1) * -1 & " nMovImporte"
    End If
    If DCtaContDem <> "" And HCtaContDem <> "" Then
        sSQL = sSQL + " UNION "
        sSQL = sSQL + "SELECT 3 nMovItem, '" & DCtaContDem & "' cCtaContCod, ISNULL(dbo.GetCtaContDesc('" & DCtaContDem & "',2,1),'') cCtaContDesc, " & nMovImporteDem * IIf(pnMoneda = gMonedaExtranjera, lnCambioFijo, 1) & " nMovImporte"
        sSQL = sSQL + " UNION "
        sSQL = sSQL + "SELECT 4 nMovItem, '" & HCtaContDem & "' cCtaContCod, ISNULL(dbo.GetCtaContDesc('" & HCtaContDem & "',2,1),'') cCtaContDesc, " & nMovImporteDem * IIf(pnMoneda = gMonedaExtranjera, lnCambioFijo, 1) * -1 & " nMovImporte"
    End If
    
    Set rs = oConect.CargaRecordSet(sSQL)
    If rs.EOF Then
       'MsgBox "No existen datos de Movimiento ", vbInformation, "Error"
       Exit Function
    End If
    nTot = 0
    nTotH = 0
    sAsiento = sAsiento + sCabecera
    nLi = nLi + 4
    sAsiento = sAsiento + BON & " EN MONEDA NACIONAL " + oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento + "--------------------" & BOFF + oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento & oImpresora.gPrnSaltoLinea
    Do While Not rs.EOF
       DoEvents
       If nLi > pnLinPage - 6 Then
           nLi = 5
           sAsiento = sAsiento + oImpresora.gPrnSaltoPagina & sCabecera
       End If
       sAsiento = sAsiento + CON & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " & Justifica(rs!cCtaContDesc, 71) & " " _
            & Right(Space(16) & IIf(rs!nMovImporte > 0, Format(rs!nMovImporte, "##,###,##0.00"), ""), 16) & " " _
            & Right(Space(16) & IIf(rs!nMovImporte < 0, Format(rs!nMovImporte * -1, "##,###,##0.00"), ""), 16) _
            & COFF & oImpresora.gPrnSaltoLinea
       If rs!nMovImporte > 0 Then
          nTot = nTot + Val(rs!nMovImporte)
       Else
          nTotH = nTotH + Val(rs!nMovImporte) * -1
       End If
       nLi = nLi + 1
       rs.MoveNext
    Loop
    sAsiento = sAsiento + CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF & oImpresora.gPrnSaltoLinea
    nLi = nLi + 2

    'JACA 20111227 Modificado para obtener cCtaContDesc Historica**************************
    'sSql = "SELECT mc.nMovItem, mc.cCtaContCod," + IIf(pbHistoCtaCont = False, " LTRIM(RTRIM(ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),'')))", " LTRIM(RTRIM(ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),'')))") + " cCtaContDesc, me.nmovMEimporte " _
    '     & "FROM   mov m join MovCta mc on mc.nmovnro=m.nmovnro JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
    '     & "WHERE  m.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
'    sSQL = "SELECT mc.nMovItem, mc.cCtaContCod," + IIf(pbHistoCtaCont = False, " LTRIM(RTRIM(ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),'')))", " LTRIM(RTRIM(ISNULL(dbo.Histo_GetCtaContDesc(mc.cCtaContCod,2,1),'')))") + " cCtaContDesc, me.nmovMEimporte " _
'         & "FROM   mov m join MovCta mc on mc.nmovnro=m.nmovnro JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'         & "WHERE  m.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
    If pnMoneda = gMonedaExtranjera Then
        If DCtaCont <> "" And HCtaCont <> "" Then
            sSQL = "SELECT 1 nMovItem, '" & DCtaCont & "' cCtaContCod, ISNULL(dbo.GetCtaContDesc('" & DCtaCont & "',2,1),'') cCtaContDesc, " & nMovImporte & " nMovMEImporte"
            sSQL = sSQL + " UNION "
            sSQL = sSQL + "SELECT 2 nMovItem, '" & HCtaCont & "' cCtaContCod, ISNULL(dbo.GetCtaContDesc('" & HCtaCont & "',2,1),'') cCtaContDesc, " & nMovImporte * -1 & " nMovMEImporte"
        End If
        If DCtaContDem <> "" And HCtaContDem <> "" Then
            sSQL = sSQL + " UNION "
            sSQL = sSQL + "SELECT 3 nMovItem, '" & DCtaContDem & "' cCtaContCod, ISNULL(dbo.GetCtaContDesc('" & DCtaContDem & "',2,1),'') cCtaContDesc, " & nMovImporteDem & " nMovMEImporte"
            sSQL = sSQL + " UNION "
            sSQL = sSQL + "SELECT 4 nMovItem, '" & HCtaContDem & "' cCtaContCod, ISNULL(dbo.GetCtaContDesc('" & HCtaContDem & "',2,1),'') cCtaContDesc, " & nMovImporteDem * -1 & " nMovMEImporte"
        End If
        
        Set rs = oConect.CargaRecordSet(sSQL)
        'JACA END**************************************************************************
        sTexto = ""
        nTot = 0
        nTotH = 0
        If Not rs.EOF Then
           sAsiento = sAsiento + BON & " EN MONEDA EXTRANJERA " & oImpresora.gPrnSaltoLinea
           sAsiento = sAsiento + "----------------------" & BOFF & oImpresora.gPrnSaltoLinea
           nLi = nLi + 2
        End If
        Do While Not rs.EOF
           DoEvents
           If nLi > pnLinPage - 6 Then
              nLi = 5
              sTexto = sTexto + oImpresora.gPrnSaltoPagina
              sTexto = sTexto + sCabecera
           End If
           sTexto = sTexto + CON & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " _
                  & Justifica(rs!cCtaContDesc, 71) & " " _
                  & IIf(rs!nMovMEImporte > 0, PrnVal(rs!nMovMEImporte, 16, 2), Space(16)) & " " _
                  & IIf(rs!nMovMEImporte < 0, PrnVal(rs!nMovMEImporte * -1, 16, 2), Space(16)) _
                  & COFF & oImpresora.gPrnSaltoLinea
           If rs!nMovMEImporte > 0 Then
              nTot = nTot + Val(rs!nMovMEImporte)
           Else
              nTotH = nTotH + Val(rs!nMovMEImporte) * -1
           End If
           nLi = nLi + 1
           rs.MoveNext
        Loop
    
        If nTot > 0 Or nTotH > 0 Then
           sAsiento = sAsiento & sTexto
           sAsiento = sAsiento + CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
           sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF + oImpresora.gPrnSaltoLinea
           nLi = nLi + 3
        End If
        If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
    End If
    If lsOpeCod <> 421113 Then
        sAsiento = sAsiento + CON & "====================================================================================================================================" & COFF & oImpresora.gPrnSaltoLinea
        
        sAsiento = sAsiento + ImprePiePag(pnColPage, psPiePag, psOtraFirma)
    End If
    sAsiento = sAsiento + "   " + oImpresora.gPrnSaltoLinea

    Set oDGeneral = Nothing
    Set oConect = Nothing
    VistaPreviaAsientoContable = sAsiento
    
    Exit Function
VistaPreviaAsientoContableErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:VistaPreviaAsientoContable Method")
End Function
'** End Juez *****************************************************************
'ALPA20140529*****************
Public Sub GeneraReporteEnArchivoExcelInicio(ByVal psNomCmac As String, ByVal psNomAge As String, ByVal psCodUser As String, ByVal pdFecsis As Date, ByVal psTitulo As String, ByVal psSubTitulo As String, _
                                    ByVal psNomArchivo As String, ByVal pMatCabeceras As Variant, ByVal prRegistros As ADODB.Recordset, _
                                    Optional pnNumDecimales As Integer, Optional Visible As Boolean = False, Optional psNomHoja As String = "", _
                                    Optional pbSinFormatDeReg As Boolean = False, _
                                    Optional pbUsarCabecerasDeRS As Boolean = False)
    Dim rs As ADODB.Recordset
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim liLineas As Integer, I As Integer
    Dim fs As Scripting.FileSystemObject
    Dim lnNumColumns As Integer

    Dim lnTotalAprobado As Double
    Dim lnTotalSaldok As Double
    Dim lnTotalDevengado As Double
    Dim lnTotalSuspenso As Double
    Dim lnTotalProvision As Double
    
    lnTotalAprobado = 0
    lnTotalSaldok = 0
    lnTotalDevengado = 0
    lnTotalSuspenso = 0
    lnTotalProvision = 0
    
    If Not (prRegistros.EOF And prRegistros.BOF) Then
        If pbUsarCabecerasDeRS = True Then
            lnNumColumns = prRegistros.Fields.Count
        Else
            lnNumColumns = UBound(pMatCabeceras)
            lnNumColumns = IIf(prRegistros.Fields.Count < lnNumColumns, prRegistros.Fields.Count, prRegistros.Fields.Count)
        End If

        If psNomHoja = "" Then psNomHoja = psNomArchivo
        psNomArchivo = psNomArchivo & "_" & psCodUser & Format(pdFecsis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"

        Set fs = New Scripting.FileSystemObject
        Set xlAplicacion = New Excel.Application
        If fs.FileExists(App.Path & "\Spooler\" & psNomArchivo) Then
            fs.DeleteFile (App.Path & "\Spooler\" & psNomArchivo)
        End If
        Set xlLibro = xlAplicacion.Workbooks.Add
        Set xlHoja1 = xlLibro.Worksheets.Add

        xlHoja1.Name = psNomHoja
        xlHoja1.Cells.Select
        
        'Cabeceras
        xlHoja1.Cells(1, 1) = psNomCmac
        xlHoja1.Cells(1, lnNumColumns) = Trim(Format(pdFecsis, "dd/mm/yyyy"))
        xlHoja1.Cells(2, 1) = psNomAge
        xlHoja1.Cells(2, lnNumColumns) = psCodUser
        xlHoja1.Cells(4, 1) = psTitulo
        xlHoja1.Cells(5, 1) = psSubTitulo
        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, lnNumColumns)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(5, lnNumColumns)).HorizontalAlignment = xlCenter

        liLineas = 6
        If pbUsarCabecerasDeRS = True Then
            For I = 0 To prRegistros.Fields.Count - 1
                xlHoja1.Cells(liLineas, I + 1) = prRegistros.Fields(I).Name
            Next I
        Else
            For I = 0 To lnNumColumns - 1
                If (I + 1) > UBound(pMatCabeceras) Then
                    xlHoja1.Cells(liLineas, I + 1) = prRegistros.Fields(I).Name
                Else
                    xlHoja1.Cells(liLineas, I + 1) = pMatCabeceras(I, 0)
                End If
            Next I
        End If
        
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).Cells.Interior.Color = RGB(220, 220, 220)
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).HorizontalAlignment = xlCenter

        If pbSinFormatDeReg = False Then
            liLineas = liLineas + 1
            While Not prRegistros.EOF
                For I = 0 To lnNumColumns - 1
                    If pMatCabeceras(I, 1) = "" Then  'Verificamos si tiene tipo
                        xlHoja1.Cells(liLineas, I + 1) = prRegistros(I)
                    Else
                        Select Case pMatCabeceras(I, 1)
                            Case "S"
                                xlHoja1.Cells(liLineas, I + 1) = prRegistros(I)
                            Case "N"
                                xlHoja1.Cells(liLineas, I + 1) = Format(prRegistros(I), "#0.00")
                            Case "D"
                                xlHoja1.Cells(liLineas, I + 1) = IIf(Format(prRegistros(I), "yyyymmdd") = "19000101", "", Format(prRegistros(I), "dd/mm/yyyy"))
                        End Select
                    End If
                Next I
                liLineas = liLineas + 1
                prRegistros.MoveNext
            Wend
        Else
            xlHoja1.Range("A7").CopyFromRecordset prRegistros 'Copia el contenido del recordset a excel
        End If
        
        xlHoja1.Range(xlHoja1.Cells(prRegistros.RecordCount + 8, 15), xlHoja1.Cells(prRegistros.RecordCount + 8, 21)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(prRegistros.RecordCount + 8, 15), xlHoja1.Cells(prRegistros.RecordCount + 8, 21)).ColumnWidth = 16
       
        xlHoja1.SaveAs App.Path & "\Spooler\" & psNomArchivo
        MsgBox "Se ha generado el Archivo en " & App.Path & "\Spooler\" & psNomArchivo

        xlAplicacion.Visible = True
        xlAplicacion.Windows(1).Visible = True


        'xlLibro.Close
        'xlAplicacion.Quit
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing

    End If
End Sub

''*** PEAC 20130612
'Public Function CargaCOAparaPVS(psPeriodo As String, psCtaIgv As String, psOpeCod As String, Optional psDocTpo As String = "") As Recordset
'    Dim oConect As DConecta
'    Dim sSql As String
'    Set oConect = New DConecta
'    sSql = "exec stp_sel_ReportesTransferenciaComprobantesParaCOA '" & psPeriodo & "','" & psCtaIgv & "','" & gPersIdRUC & "','" & gMovEstContabMovContable & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagExtornado & "','" & gMovFlagModificado & "', '" & psOpeCod & "', '" & psDocTpo & "'"
'    oConect.AbreConexion
'    Set CargaCOAparaPVS = oConect.CargaRecordSet(sSql)
'    oConect.CierraConexion
'    Set oConect = Nothing
'End Function
'PASIERS014-2015***********************************************************************
Public Function ImprimeCuadroDepreciacionAdjudicados(pdFecha As Date, pnLinPage As Integer, ByRef rsDepre As ADODB.Recordset, Optional psCtaCont As String) As String
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim nLin As Integer
Dim P    As Integer
Dim sTit As String, sTitBarra As String
Dim sTexto As String
Dim nTotDif As Currency
Dim oAju  As New DAjusteCont
Dim nTotE As Currency, nTotC As Currency
Dim lsError As String
Dim lsCuentaCont As String
Dim lsCtaDesc As String * 70
Dim loAsiento As DCtaCont
On Error GoTo ImprimeCuadroDeprecAdjudicaErr

sTit = "D E P R E C I A C I Ó N  D E  A D J U D I C A D O S"
sTitBarra = "DEPRECIACIÓN DE ADJUDICADOS"
RaiseEvent BarraShow(1)
RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo Datos", vbBlue)
Set rsDepre = oAju.AjusteDepreciacionAdjudicados(Format(pdFecha, "yyyyMMdd"))
RaiseEvent BarraClose
If rsDepre.EOF Then
    MsgBox "No existen diferencias entre Estadísticas y Saldos Contables", vbInformation, "Aviso"
    ImprimeCuadroDepreciacionAdjudicados = ""
    Exit Function
Else
     Set loAsiento = New DCtaCont
        nLin = pnLinPage
        P = 0
   RaiseEvent BarraShow(rsDepre.RecordCount)
   Do While Not rsDepre.EOF
        RaiseEvent BarraProgress(rsDepre.Bookmark, sTitBarra, "", "Generando Cuadro...", vbBlue)
        If nLin > pnLinPage - 6 Then
            Linea sTexto, Cabecera(sTit, P, "", 120, , , gsNomCmac), 0
            'Modificado PASI20160213
            'Linea sTexto, "Cuenta Contable             S.Estadistico      "
            'Linea sTexto, "-------------------------------------------------------"
            Linea sTexto, "Cuenta Contable" & Space(20) & "Descripción" & Space(85) & "DEBE" & Space(10) & "HABER"
            Linea sTexto, "----------------------------------------------------------------------------------------------------------------------------------------------------------"
            
            'end PASI***
            nLin = 5
        End If
        gnImporte = nVal(Format(rsDepre!nSaldo, "#.00"))
        lsCuentaCont = rsDepre!CtaDebe
        On Error Resume Next
        lsError = loAsiento.ExisteCuentaCad(Trim(lsCuentaCont), True)
        If lsError <> "" Then
            Linea sTexto, loAsiento.ExisteCuentaCad(Trim(lsCuentaCont), True) & lsCuentaCont
        End If
        On Error GoTo ImprimeCuadroDeprecAdjudicaErr
        lsCtaDesc = rsDepre!CtaDebeDesc
        Linea sTexto, Justifica(rsDepre!CtaDebe, 22) & Space(12) & lsCtaDesc & Space(20) & PrnVal(gnImporte, 14, 2)
        
        'PASI20160213
        lsCuentaCont = rsDepre!CtaHaber
        On Error Resume Next
        lsError = loAsiento.ExisteCuentaCad(Trim(lsCuentaCont), True)
        If lsError <> "" Then
            Linea sTexto, loAsiento.ExisteCuentaCad(Trim(lsCuentaCont), True) & lsCuentaCont
        End If
        On Error GoTo ImprimeCuadroDeprecAdjudicaErr
        lsCtaDesc = rsDepre!CtaHaberDesc
        Linea sTexto, Justifica(rsDepre!CtaHaber, 22) & Space(12) & lsCtaDesc & Space(35) & PrnVal(gnImporte, 14, 2)
        'end PASI
        
        nTotE = nTotE + gnImporte
        nLin = nLin + 1
        rsDepre.MoveNext
    Loop
    Linea sTexto, "----------------------------------------------------------------------------------------------------------------------------------------------------------"
    'Linea sTexto, "    TOTAL                   " & PrnVal(nTotE, 14, 2) & "           " & PrnVal(nTotE, 14, 2) 'Comentado PASI20160215
    Linea sTexto, "    TOTAL      " & Space(109) & PrnVal(nTotE, 14, 2) & Space(1) & PrnVal(nTotE, 14, 2) 'PASI20160215
    
    rsDepre.MoveFirst
    ImprimeCuadroDepreciacionAdjudicados = sTexto
    RaiseEvent BarraClose
    Set loAsiento = Nothing
End If
Exit Function
ImprimeCuadroDeprecAdjudicaErr:
    RaiseEvent BarraClose
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeCuadroDepreciacionAdjudicados Method")
End Function
'END PASI
'PASI20160216 ERS0762015
Public Function ImprimeCuadroAjusteParaGramosDeOroEnCustodia(pdFecha As Date, pnLinPage As Integer, ByRef rsAjuste As ADODB.Recordset) As String
Dim nLin As Integer
Dim P As Integer
Dim sTit As String, sTitBarra As String
Dim sTexto As String
Dim nTotDif As Currency
Dim oAju As New DAjusteCont
Dim nTotE As Currency
Dim lsError As String
Dim lsCtaCont As String
Dim lsCtaDesc As String * 70
Dim loAsiento As DCtaCont
On Error GoTo ImprimeCuadroAjusteParaGramosDeOroEnCustodiaErr
sTit = "A J U S T E  D E  G R A M O S  D E  O R O  E N  C U S T O D I A"
sTitBarra = "AJUSTE GRAMOS DE ORO EN CUSTODIA"
RaiseEvent BarraShow(1)
RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo Datos", vbBlue)
Set rsAjuste = oAju.AjusteParaGramosDeOroEnCustodia(Format(pdFecha, "yyyyMMdd"))
RaiseEvent BarraClose
If rsAjuste.EOF Then
    MsgBox "Por el momento no existe información para generar el Ajuste de Gramos de Oro en Custodia"
    ImprimeCuadroAjusteParaGramosDeOroEnCustodia = ""
Else
    Set loAsiento = New DCtaCont
        nLin = pnLinPage
        P = 0
    RaiseEvent BarraShow(rsAjuste.RecordCount)
    Do While Not rsAjuste.EOF
        RaiseEvent BarraProgress(rsAjuste.Bookmark, sTitBarra, "", "Generando Cuadro...", vbBlue)
        If nLin > pnLinPage - 6 Then
            Linea sTexto, Cabecera(sTit, P, "", 120, , , gsNomCmac), 0
            Linea sTexto, "Cuenta Contable" & Space(20) & "Descripción" & Space(80) & "DEBE" & Space(10) & "HABER"
            Linea sTexto, "----------------------------------------------------------------------------------------------------------------------------------------------------------"
            nLin = 5
        End If
        lsCtaCont = rsAjuste!cCtaContCod
        On Error Resume Next
        lsError = loAsiento.ExisteCuentaCad(Trim(lsCtaCont), True)
        If lsError <> "" Then
            Linea sTexto, loAsiento.ExisteCuentaCad(Trim(lsCtaCont), True) & lsCtaCont
        End If
        On Error GoTo ImprimeCuadroAjusteParaGramosDeOroEnCustodiaErr
        lsCtaDesc = rsAjuste!cCtaContDesc
        If rsAjuste!nSaldoDebe > 0 Then
            gnImporte = nVal(Format(rsAjuste!nSaldoDebe, "#.00"))
            Linea sTexto, Justifica(rsAjuste!cCtaContCod, 22) & Space(12) & lsCtaDesc & Space(20) & Abs(PrnVal(gnImporte, 14, 2))
        Else
            gnImporte = nVal(Format(rsAjuste!nSaldoHaber, "#.00"))
            Linea sTexto, Justifica(rsAjuste!cCtaContCod, 22) & Space(12) & lsCtaDesc & Space(35) & Abs(PrnVal(gnImporte, 14, 2))
        End If
        gnImporte = rsAjuste!nSaldoDebe
        nTotE = nTotE + gnImporte
        nLin = nLin + 1
        rsAjuste.MoveNext
    Loop
    Linea sTexto, "----------------------------------------------------------------------------------------------------------------------------------------------------------"
    Linea sTexto, "    TOTAL      " & Space(109) & PrnVal(nTotE, 14, 2) & Space(1) & PrnVal(nTotE, 14, 2)
    
    rsAjuste.MoveFirst
    ImprimeCuadroAjusteParaGramosDeOroEnCustodia = sTexto
    RaiseEvent BarraClose
    Set loAsiento = Nothing
End If
Exit Function
ImprimeCuadroAjusteParaGramosDeOroEnCustodiaErr:
    RaiseEvent BarraClose
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeCuadroAjusteParaGramosDeOroEnCustodia Method")
End Function
'END PASI

Public Function ImprimeCuadroOtrosProcesos(ByVal psAsiento As String, ByVal rs As ADODB.Recordset, ByVal sTitBarra As String, ByVal sTit As String, ByVal pnLinPage As Integer, ByVal loAsiento As DCtaCont) As String
'NAGL 202102 Agregó psAsiento As String
Dim nLin As Integer
Dim P As Integer
Dim sTexto As String, sTitTipo As String, sTitDet As String
Dim nTotDif As Currency
Dim nTotE As Currency, nTotC As Currency
Dim lsError As String
Dim lsCuentaCont As String
Dim nReg As Integer
'Dim loAsiento As New DCtaCont
Dim psTipoRep As String, psSeccion As String
psTipoRep = ""
psSeccion = "SC"
sTitTipo = ""
sTitDet = ""
nLin = pnLinPage
P = 0
nReg = 1
RaiseEvent BarraShow(rs.RecordCount)
Do While Not rs.EOF
   RaiseEvent BarraProgress(rs.Bookmark, sTitBarra, "", "Generando Cuadro...", vbBlue)
   If nReg = 1 Then
      psTipoRep = rs!cTipo
   End If
   If (psSeccion <> rs!cSeccion) Then
      If psAsiento = "34" Then
        Linea sTexto, "--------------------------------------------------------------------------------"
        Linea sTexto, "TOTAL DIFERENCIA    " & PrnVal(nTotE, 18, 2) & "   " & PrnVal(nTotC, 18, 2) & "   " & PrnVal(nTotDif, 18, 2)
      Else
        Linea sTexto, "-------------------------------------------------------------------------------------"
        Linea sTexto, "TOTAL DIFERENCIA           " & PrnVal(nTotE, 18, 2) & "  " & PrnVal(nTotC, 18, 2) & "  " & PrnVal(nTotDif, 18, 2)
      End If 'NAGL 202102
      nTotDif = 0
      nTotE = 0
      nTotC = 0
      nLin = pnLinPage
   End If
   If nLin > pnLinPage - 6 Then
      If psTipoRep <> rs!cTipo Or Right(rs!cTipo, 1) = "1" Then
         sTitTipo = " " & Right(rs!cTipo, 1)
      End If
      If rs!cSeccion = "SC" Then
         sTitDet = " - CAPITAL"
      ElseIf rs!cSeccion = "ID" Or rs!cSeccion = "IVD" Then
         sTitDet = " - INTERES DEVENGADO"
      ElseIf rs!cSeccion = "IP" Then
         sTitDet = " - INTERES PERCIBIDO"
      End If
      If nReg = 1 Then
         Linea sTexto, Cabecera("", P, "", 110, , , gsNomCmac), 0
      Else
         Linea sTexto, ""
         Linea sTexto, ""
      End If
      If psAsiento = "34" Then
         Linea sTexto, Centra(sTit & sTitTipo & sTitDet)
         Linea sTexto, "Cuenta Contable          S.Estadistico           S.Contable           Diferencia"
         Linea sTexto, "--------------------------------------------------------------------------------"
      Else
        Linea sTexto, Centra(sTit & sTitTipo & sTitDet)
        Linea sTexto, "Cuenta Contable                 S.Estadistico          S.Contable          Diferencia"
        Linea sTexto, "-------------------------------------------------------------------------------------"
      End If 'NAGL 202102
      nLin = 5
   End If
   gnImporte = nVal(Format(rs!nSaldo, "#.00"))
   lsCuentaCont = IIf(IsNull(rs!Cta1), rs!cta2, rs!Cta1)
   lsError = loAsiento.ExisteCuentaCad(Trim(lsCuentaCont), True)
   If lsError <> "" Then
     Linea sTexto, loAsiento.ExisteCuentaCad(Trim(lsCuentaCont), True) & lsCuentaCont
   End If
   
   If psAsiento = "34" Then
        Linea sTexto, Justifica(IIf(IsNull(rs!Cta1), rs!cta2, rs!Cta1), 14) & "      " & PrnVal(gnImporte, 18, 2) & "   " & PrnVal(rs!nCtaSaldoImporte, 18, 2) & "   " & PrnVal(rs!nCtaSaldoImporte - rs!nSaldo, 18, 2)
   Else
        Linea sTexto, Justifica(IIf(IsNull(rs!Cta1), rs!cta2, rs!Cta1), 22) & "     " & PrnVal(gnImporte, 18, 2) & "  " & PrnVal(rs!nCtaSaldoImporte, 18, 2) & "  " & PrnVal(rs!nCtaSaldoImporte - rs!nSaldo, 18, 2)
   End If 'NAGL 202102
   
   nTotDif = nTotDif + (rs!nCtaSaldoImporte - gnImporte)
   nTotE = nTotE + gnImporte
   nTotC = nTotC + rs!nCtaSaldoImporte
   nLin = nLin + 1
   If nReg = rs.RecordCount Then
      If psAsiento = "34" Then
        Linea sTexto, "--------------------------------------------------------------------------------"
        Linea sTexto, "TOTAL DIFERENCIA    " & PrnVal(nTotE, 18, 2) & "   " & PrnVal(nTotC, 18, 2) & "   " & PrnVal(nTotDif, 18, 2)
      Else
        Linea sTexto, "-------------------------------------------------------------------------------------"
        Linea sTexto, "TOTAL DIFERENCIA           " & PrnVal(nTotE, 18, 2) & "  " & PrnVal(nTotC, 18, 2) & "  " & PrnVal(nTotDif, 18, 2)
      End If 'NAGL 202102
      nTotDif = 0
      nTotE = 0
      nTotC = 0
   End If
   psSeccion = rs!cSeccion
   psTipoRep = rs!cTipo
   nReg = nReg + 1
   rs.MoveNext
Loop
ImprimeCuadroOtrosProcesos = sTexto
End Function 'NAGL 20200804 Según ACTA N°049-2020

Public Function ImprimeCuentasObservacion(ByVal rs As ADODB.Recordset, ByVal pnTipoAsiento As Integer)
Dim sTexto As String
Dim nReg As Integer, P As Integer
nReg = 1
P = 0
  If pnTipoAsiento = 32 Then
    Do While Not rs.EOF
        If nReg = 1 Then
            Linea sTexto, Cabecera("", P, "", 110, , , gsNomCmac), 0
            Linea sTexto, "  OBSERVACIÓN CONTABLE DE INTERESES DIFERIDOS EN CRÉDITOS CANCELADOS"
            Linea sTexto, "Observación                   Cuenta Contable              Crédito"
            Linea sTexto, "-----------------------------------------------------------------------"
        End If
        Linea sTexto, Justifica(IIf(IsNull(rs!cObservacion), "", rs!cObservacion), 25) & "     " & Justifica(IIf(IsNull(rs!cCtaCnt), "", rs!cCtaCnt), 22) & "  " & Justifica(IIf(IsNull(rs!cCtaCod), "", rs!cCtaCod), 22)
        nReg = nReg + 1
        rs.MoveNext
    Loop
  End If
  ImprimeCuentasObservacion = sTexto
End Function 'NAGL 202008 Según ACTA N°063-2020
