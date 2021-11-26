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
Dim vsConexion As String
Dim vsBaseComunes As String
Dim vsBasePesonas As String
Type tBilletajes
    lsBilletes As Currency
    lnMontoBil As Currency
    lsMonedas As Currency
    lnMontoMon As Currency
End Type

Dim vsAgeNom As String
Dim vsEmpNom As String
Dim vsFecSis As String

Public Event BarraShow(pnMax)
Public Event BarraProgress(value, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
Public Event BarraClose()

'##ModelId=3A80BB5102DE
Public Sub ImprimeGeneral(pnTipoRep As Integer)
    On Error GoTo ImprimeGeneralErr

    'your code goes here...

    Exit Sub
ImprimeGeneralErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeGeneral Method")
End Sub

'##ModelId=3A836E54008C
Public Function ImprimeAsientoContable(ByVal psMovNro As String, ByVal pnLinPage As Integer, ByVal pnColPage As Integer, Optional ByVal psTit1 As String, Optional ByVal psCabe1 As String, Optional ByVal psPiePag As String = "9", Optional ByVal psOtraFirma As String = "", Optional ByVal pbMoneda As Boolean = True) As String
    'On Error GoTo ImprimeAsientoContableErr
    Dim sSql As String
    Dim Rs       As New ADODB.Recordset
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
    
    Dim lsOrigenCab As String
    Dim lsDestinoCab As String
    
    
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
        lsFecha = CDate(Mid(psMovNro, 7, 2) + "/" + Mid(psMovNro, 5, 2) + "/" + Mid(psMovNro, 1, 4))
        
        lnCambioFijo = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCFijoMes)
        lnCambioVenta = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCVenta)
        lnCambioCompra = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCCompra)
    
       sSql = "SELECT * FROM Mov WHERE cMovNro = '" & psMovNro & "'"
       Set Rs = oConect.CargaRecordSet(sSql)
       If Rs.EOF Then
          Exit Function
       End If
       lsGlosa = Replace(Rs!cMovDesc, Chr(13) & oImpresora.gPrnSaltoLinea, " ")
       lsMovNro = psMovNro
       lsOpeCod = Rs!COPECOD
       If Rs!nMovFlag = gMovFlagDeExtorno Then
            psTit1 = " A S I E N T O  C O N T A B L E   D E   E X T O R N O"
       End If
       sAsiento = sAsiento & ImpreCabAsiento(pnColPage, CDate(lsFecha), "", lsOpeCod, lsMovNro, psTit1, pbMoneda) & oImpresora.gPrnSaltoLinea
       nLi = 5
       If psCabe1 <> "" Then
          sAsiento = sAsiento & psCabe1 & oImpresora.gPrnSaltoLinea
          nLi = nLi + 4
       Else
           'BUSCA LAS PERSONA INVOLUCRADAS EN EL MOVIMIENTO
            sSql = " SELECT Distinct  M.CMOVNRO,  " _
                    & "        " _
                    & "        " _
                    & "         P.cPersCod, " _
                    & "         Convert(Char(40),P.CPERSNOMBRE) as cPersNombre " _
                    & "         " _
                    & " FROM    MOVCTA MC JOIN MOV M    ON M.NMOVNRO = MC.NMOVNRO " _
                    & "         " _
                    & "         JOIN MOVGASTO MOI         ON MOI.NMOVNRO = MC.NMOVNRO " _
                    & "         JOIN PERSONA P          ON P.CPERSCOD = MOI.CPERSCOD " _
                    & " WHERE   M.cMovNro = '" & lsMovNro & "'"
              
            Set Rs = oConect.CargaRecordSet(sSql)
            sDestino = "": sOrigen = ""
            Do While Not Rs.EOF
                sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & PstaNombre(Rs!cPersNombre)
                nLi = nLi + 1
                Rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat("Persona(s) Origen  ", 25) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat("Persona(s) Destino ", 25) & " :" & sDestino & oImpresora.gPrnSaltoLinea
            End If
            
            sSql = " SELECT  M.CMOVNRO, " _
                & "         CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                 WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "         P.cPersCod , CI.cCtaIFCod , " _
                & "         Convert(Char(40),CTPOI.cConsDescripcion ) as Entidad, " _
                & "         Convert(Char(40),P.CPERSNOMBRE) as cPersNombre, " _
                & "         Convert(char(50),C.cConsDescripcion + ' ' + CI.cCtaIFDesc) as DescCtaIF, " _
                & "         MO.COBJETOCOD , MC.nMovImporte " _
                & " FROM    MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "         JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "         JOIN MOVOBJIF MOI       ON MOI.NMOVNRO = MO.NMOVNRO AND MOI.NMOVITEM = MO.NMOVITEM " _
                & "         JOIN CTAIF CI           ON CI.CCTAIFCOD = MOI.CCTAIFCOD AND MOI.CPERSCOD = CI.CPERSCOD AND MOI.CIFTPO = CI.CIFTPO  " _
                & "         JOIN CONSTANTE C        ON C.nConsValor = SUBSTRING(CI.CCTAIFCOD,2,1) AND C.nCONSCOD LIKE '" & gCGTipoCtaIF & "' " _
                & "         JOIN INSTITUCIONFINANC I    ON I.CPERSCOD = CI.CPERSCOD AND I.CIFTPO = CI.CIFTPO " _
                & "         JOIN PERSONA P          ON P.CPERSCOD = I.CPERSCOD " _
                & "         JOIN CONSTANTE CTPOI        ON CTPOI.nConsValor = I.cIFTpo AND CTPOI.nCONSCOD LIKE '" & gCGTipoIF & "' " _
                & " WHERE M.cMovNro = '" & lsMovNro & "' AND MO.cObjetoCod ='" & Format(ObjEntidadesFinancieras, "00") & "'"
            
            Set Rs = oConect.CargaRecordSet(sSql)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not Rs.EOF
                If Rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(Rs!cPersNombre) & " " & Trim(Rs!DescCtaIF)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(Rs!cPersNombre) & " " & Trim(Rs!DescCtaIF)
                End If
                nLi = nLi + 1
                Rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat("Cuenta(s) If Origen", 25) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat("Cuenta(s) IF Destino", 25) & " :" & sDestino & oImpresora.gPrnSaltoLinea
            End If
            
            sSql = "    SELECT  M.CMOVNRO, " _
                & "             CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                     WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "             MOA.cAreaCod + MOA.CAGECOD AS cAreaCod, A.cAreaDescripcion + ' ' + ISNULL(AG.CAGEDESCRIPCION,'' )  as cAreaDescripcion , MO.COBJETOCOD, MC.nMovImporte , O.cObjetoDesc " _
                & "     FROM    MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "             JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "             JOIN MovObjAreaAgencia MOA         ON MOA.NMOVNRO = MO.NMOVNRO AND MOA.NMOVITEM = MO.NMOVITEM " _
                & "             JOIN AREAS A            ON A.cAreaCod = MOA.cAreaCod LEFT JOIN AGENCIAS AG ON AG.CAGECOD = MOA.CAGECOD " _
                & "             JOIN OBJETO O ON O.cObjetoCod  = MO.cObjetoCod " _
                & "     WHERE   M.cMovNro = '" & lsMovNro & "' AND Mo.cObjetoCod IN ('" & Format(ObjCMACAgenciaArea, "00") & "','" & Format(ObjCMACArea, "00") & "')"

            Set Rs = oConect.CargaRecordSet(sSql)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not Rs.EOF
                lsDesObjeto = Trim(Rs!cObjetoDesc)
                If Rs!cDH = "D" Then
                    If lsDestinoCab <> IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(Rs!cAreaCod) & " - " & Trim(Rs!CAREADESCRIPCION) Then
                        lsDestinoCab = IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(Rs!cAreaCod) & " - " & Trim(Rs!CAREADESCRIPCION)
                        sDestino = sDestino & lsDestinoCab  'IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea  & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!CAREADESCRIPCION)
                    End If
                    
                    'sDestino = sDestino & lsDestinoCab  'IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea  & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!CAREADESCRIPCION)
                Else
                    If lsOrigenCab <> IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(Rs!cAreaCod) & " - " & Trim(Rs!CAREADESCRIPCION) Then
                        lsOrigenCab = IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(Rs!cAreaCod) & " - " & Trim(Rs!CAREADESCRIPCION)
                        sOrigen = sOrigen & lsOrigenCab 'IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea  & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!CAREADESCRIPCION)
                    End If
                
                
                    'sOrigen = sOrigen & lsOrigenCab 'IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea  & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!CAREADESCRIPCION)
                End If
                nLi = nLi + 1
                Rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Origen ", 25) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Destino ", 25) & " :" & sDestino & oImpresora.gPrnSaltoLinea
            End If
            
            'sSQL = " SELECT    MC.CMOVNRO, " _
                & "             CASE      WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                     WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "             MOA.cAgecod , A.cAgeDescripcion, MO.COBJETOCOD, MC.nMovImporte , O.cObjetoDesc " _
                & "   FROM      MOVCTA MC " _
                & "             JOIN MOVOBJ MO          ON MO.CMOVNRO = MC.CMOVNRO AND MC.CMOVITEM=MO.CMOVITEM " _
                & "             JOIN MovObjAreaAgencia MOA      ON MOA.CMOVNRO = MO.CMOVNRO AND MOA.CMOVITEM = MO.CMOVITEM " _
                & "             JOIN Agencias A         ON A.cAgeCod = MOA.cAgeCod " _
                & "             JOIN OBJETO O ON O.cObjetoCod  = MO.cObjetoCod " _
                & "   WHERE     MC.cMovNro='" & lsMovNro & "' AND Mo.cObjetoCod IN ('" & Format(ObjCMACAgencias, "00") & "','" & Format(ObjCMACAgenciaArea, "00") & "')"
            
            'Set rs = oConect.CargaRecordSet(sSQL)
            'sDestino = "": sOrigen = "": lsDesObjeto = ""
            'Do While Not rs.EOF
            '    lsDesObjeto = Trim(rs!cObjetoDesc)
            '    If rs!cDH = "D" Then
            '        sDestino = sDestino & IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea  & Space(14), "") & Trim(rs!CAGECOD) & " - " & Trim(rs!CAGEDESCRIPCION)
            '    Else
            '        sOrigen = sOrigen & IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea  & Space(14), "") & Trim(rs!CAGECOD) & " - " & Trim(rs!CAGEDESCRIPCION)
            '    End If
            '    nLi = nLi + 1
            '    rs.MoveNext
            'Loop
            'If sOrigen <> "" Then
            '    sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Origen ", 25) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
            'End If
            'If sDestino <> "" Then
            '    sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Destino ", 25) & " :" & sDestino & oImpresora.gPrnSaltoLinea
            'End If
            
            sSql = " SELECT    M.CMOVNRO, " _
                & "             CASE      WHEN mc.nMovImporte>0 THEN 'D' " _
                & "                     WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                & "             MO.COBJETOCOD, MC.nMovImporte, O.cObjetoDesc " _
                & "   FROM      MOVCTA MC JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
                & "             JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                & "             JOIN OBJETO O ON O.cObjetoCod  = MO.cObjetoCod " _
                & "   WHERE     M.cMovNro='" & lsMovNro & "' " _
                & "             AND MO.cObjetoCod NOT IN ('" & Format(ObjCMACAgencias, "00") & "','" & Format(ObjCMACAgenciaArea, "00") & "','" & Format(ObjDescomEfectivo, "00") & "','" & Format(ObjEntidadesFinancieras, "00") & "','" & Format(ObjBienesServicios, "00") & "','" & Format(ObjPersona, "00") & "')"
            
            Set Rs = oConect.CargaRecordSet(sSql)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not Rs.EOF
                lsDesObjeto = "OBJETOS"
                If Rs!cDH = "D" Then
                    If lsDestinoCab <> IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(Rs!cObjetoCod) & "-" & Trim(Rs!cObjetoDesc) Then
                        lsDestinoCab = IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(Rs!cObjetoCod) & "-" & Trim(Rs!cObjetoDesc)
                    Else
                        lsDestinoCab = ""
                    End If
                    
                    sDestino = sDestino & lsDestinoCab  ' IIf(Len(sDestino) > 0, oImpresora.gPrnSaltoLinea  & Space(14), "") & Trim(rs!cObjetoCod) & "-" & Trim(rs!cObjetoDesc)
                Else
                    If lsOrigenCab <> IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(Rs!cObjetoCod) & "-" & Trim(Rs!cObjetoDesc) Then
                        lsOrigenCab = IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea & Space(14), "") & Trim(Rs!cObjetoCod) & "-" & Trim(Rs!cObjetoDesc)
                    Else
                        lsOrigenCab = ""
                    End If
                    
                    sOrigen = sOrigen & lsOrigenCab  ' IIf(Len(sOrigen) > 0, oImpresora.gPrnSaltoLinea  & Space(14), "") & Trim(rs!cObjetoCod) & "-" & Trim(rs!cObjetoDesc)
                End If
                nLi = nLi + 1
                Rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Origen ", 25) & " :" & sOrigen & oImpresora.gPrnSaltoLinea
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Destino", 25) & " :" & sDestino & oImpresora.gPrnSaltoLinea
            End If
            
        End If ' del pscab1
    
       'Tipo de Cambio
       
        sSql = "SELECT nMovTpoCambio FROM MovTpoCambio MC JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE M.cMovNro = '" & lsMovNro & "'"
        Set Rs = oConect.CargaRecordSet(sSql)
        If Not Rs.EOF Then
            sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "      T.C.Mercado: " & Format(Rs!nMovTpoCambio, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
        Else
            If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Then
                sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "     T.C.Compra : " & Format(lnCambioCompra, "##,###,##0.000") & "     T.C.Venta : " & Format(lnCambioVenta, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
            End If
        End If
       nLi = nLi + 1
       'Datos de Documentos
       sSql = "SELECT b.cDocAbrev, a.cDocNro, a.dDocFecha FROM MovDoc a join mov m on m.nmovnro =a.nmovnro  , " & vsBaseComunes & "Documento b " _
            & "WHERE m.cMovNro = '" & lsMovNro & "' and b.nDocTpo = a.nDocTpo "
       
        Set rsDoc = oConect.CargaRecordSet(sSql)
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
          sAsiento = sAsiento + ImpreFormat("Documentos", 25) & " :" & sDoc & oImpresora.gPrnSaltoLinea
        End If
        rsDoc.Close
       'Glosa
       sAsiento = sAsiento & ImpreGlosa(lsGlosa, pnColPage) + oImpresora.gPrnSaltoLinea
       nLi = nLi + 2
       lsMovBilletaje = ""
       lsMovBilletaje = ImprimeBilletaje(Mid(lsOpeCod, 3, 1), lsMovNro, pnColPage)
       If lsMovBilletaje <> "" Then
            sAsiento = sAsiento + lsMovBilletaje & oImpresora.gPrnSaltoLinea
            nLi = nLi + 8
       End If
    'Cabecera 2 de Asiento
    sCabecera = sCabecera + CON
    sCabecera = sCabecera + "====================================================================================================================================" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "Itm C U E N T A   C O N T A B L E                                                                           DEBE            HABER" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "    Codigo              Descripcion" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "------------------------------------------------------------------------------------------------------------------------------------" & COFF & oImpresora.gPrnSaltoLinea
    'MovCta para MN
    sSql = "SELECT mc.nMovItem, mc.cCtaContCod, mc.nMovImporte " _
         & "FROM   MovCta mc JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE  M.cMovNro = '" & lsMovNro & "' "
    Set Rs = oConect.CargaRecordSet(sSql)
    If Rs.EOF Then
       'MsgBox "No existen datos de Movimiento ", vbInformation, "Error"
       Exit Function
    End If
    nTot = 0
    nTotH = 0
    sAsiento = sAsiento + sCabecera
    nLi = nLi + 4
    sAsiento = sAsiento + BON & " EN MONEDA NACIONAL " + oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento + "--------------------" & BOFF + oImpresora.gPrnSaltoLinea
    Do While Not Rs.EOF
      'DoEvents
       If nLi > pnLinPage - 6 Then
           nLi = 5
           sAsiento = sAsiento + oImpresora.gPrnSaltoPagina & sCabecera
       End If
       sAsiento = sAsiento + CON & Format(Rs!nMovItem, "000") & " " & Justifica(Rs!cCtaContCod, 20) & " " & Justifica(oDGeneral.CuentaNombre(Rs!cCtaContCod), 71) & " " _
            & Right(Space(16) & IIf(Rs!NMOVIMPORTE > 0, Format(Rs!NMOVIMPORTE, "##,###,##0.00"), ""), 16) & " " _
            & Right(Space(16) & IIf(Rs!NMOVIMPORTE < 0, Format(Rs!NMOVIMPORTE * -1, "##,###,##0.00"), ""), 16) _
            & COFF & oImpresora.gPrnSaltoLinea
       If Rs!NMOVIMPORTE > 0 Then
          nTot = nTot + Val(Rs!NMOVIMPORTE)
       Else
          nTotH = nTotH + Val(Rs!NMOVIMPORTE) * -1
       End If
       nLi = nLi + 1
       Rs.MoveNext
    Loop
    sAsiento = sAsiento + CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF & oImpresora.gPrnSaltoLinea
    nLi = nLi + 2
    
    sSql = "SELECT mc.nMovItem, mc.cCtaContCod, me.nmovMEimporte " _
         & "FROM   mov m join MovCta mc on mc.nmovnro=m.nmovnro JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
         & "WHERE  m.cMovNro = '" & lsMovNro & "'"
    Set Rs = oConect.CargaRecordSet(sSql)
    sTexto = ""
    nTot = 0
    nTotH = 0
    If Not Rs.EOF Then
       sAsiento = sAsiento + BON & " EN MONEDA EXTRANJERA " & oImpresora.gPrnSaltoLinea
       sAsiento = sAsiento + "----------------------" & BOFF & oImpresora.gPrnSaltoLinea
       nLi = nLi + 2
    End If
    Do While Not Rs.EOF
      'DoEvents
       If nLi > pnLinPage - 6 Then
          nLi = 5
          sTexto = sTexto + oImpresora.gPrnSaltoPagina
          sTexto = sTexto + sCabecera
       End If
       sTexto = sTexto + CON & Format(Rs!nMovItem, , "000") & " " & Justifica(Rs!cCtaContCod, 20) & " " _
              & Justifica(oDGeneral.CuentaNombre(Rs!cCtaContCod), 71) & " " _
              & IIf(Rs!NMOVMEIMPORTE > 0, PrnVal(Rs!NMOVMEIMPORTE, 16, 2), Space(16)) & " " _
              & IIf(Rs!NMOVMEIMPORTE < 0, PrnVal(Rs!NMOVMEIMPORTE * -1, 16, 2), Space(16)) _
              & COFF & oImpresora.gPrnSaltoLinea
       If Rs!NMOVMEIMPORTE > 0 Then
          nTot = nTot + Val(Rs!NMOVMEIMPORTE)
       Else
          nTotH = nTotH + Val(Rs!NMOVMEIMPORTE) * -1
       End If
       nLi = nLi + 1
       Rs.MoveNext
    Loop
    If nTot > 0 Or nTotH > 0 Then
       sAsiento = sAsiento & sTexto
       sAsiento = sAsiento + CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
       sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF + oImpresora.gPrnSaltoLinea
       nLi = nLi + 3
    End If
    If Rs.State = adStateOpen Then Rs.Close: Set Rs = Nothing
    sAsiento = sAsiento + CON & "====================================================================================================================================" & COFF & oImpresora.gPrnSaltoLinea
    
    sAsiento = sAsiento + ImprePiePag(pnColPage, psPiePag, psOtraFirma)
    
    sAsiento = sAsiento + "   " + oImpresora.gPrnSaltoLinea

    Set oDGeneral = Nothing
    Set oConect = Nothing
    ImprimeAsientoContable = sAsiento
    
    Exit Function
ImprimeAsientoContableErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeAsientoContable Method")
End Function

'##ModelId=3A8D63BE03C8
Public Function ImprimeReciboEgresos(ByVal pnColPage As Integer, ByVal psMovNro As String, ByVal psGlosa As String, _
                                ByVal pdFecha As String, ByVal psEmpresaLogo As String, ByVal psOpeCod As String, _
                                ByVal lbEsCajaChica As Boolean, ByVal psDescCajaChica As String, ByVal pnArendirFase As ARendirFases, _
                                ByVal psDocNroARendir As String, ByVal psDocNro As String, ByVal psFechaDoc As String, _
                                ByVal psPersCodProv As String, ByVal psPersNombreProv As String, _
                                ByVal psCargo As String, ByVal pnTotal As Currency) As String
    Dim sTexto As String
    Dim nAncho As Integer, n As Integer
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
  
    lsTexto = lsTexto + ImpreCabAsiento(pnColPage, pdFecha, psEmpresaLogo, psOpeCod, psMovNro, " R E C I B O   D E   E G R E S O  Nro. " & psDocNro, , , , 1) & oImpresora.gPrnSaltoLinea
    If lbEsCajaChica Then
         lsTexto = lsTexto + "  CAJA CHICA   : " & BON & Mid(psDescCajaChica, 1, 60) & BOFF & oImpresora.gPrnSaltoLinea
    End If
    If pnArendirFase = ArendirRendicion Or pnArendirFase = ArendirSustentacion Then
        lsTexto = lsTexto + "  Rec. Arendir : " & BON & psDocNroARendir & BOFF & "  Fecha : " & BON & psFechaDoc & BOFF & oImpresora.gPrnSaltoLinea
    End If
    lsTexto = lsTexto + "  Persona      : " & BON & psPersCodProv & "  " & ImpreCarEsp(psPersNombreProv) & BOFF & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + "  Cargo        : " & psCargo & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + "  Importe      : " & BON & ConvNumLet(Trim(Str(pnTotal)), False) & BOFF & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + ImpreGlosa(psGlosa, pnColPage, "  Concepto     : ") & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + "" + oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + BON & Space(15) & "_________________________                ______________________" & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + Space(15) & CON & Centra(ImpreCarEsp(Mid(psPersNombreProv, 1, 40)), 40) & COFF & "                    Vo Bo JEFE AREA " & BOFF & oImpresora.gPrnSaltoLinea
  
    ImprimeReciboEgresos = lsTexto
    
    Exit Function
ImprimeReciboEgresosErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeReciboEgresos Method")
End Function

'##ModelId=3A969A8500EA
Public Function ImprimeReciboARendir(ByVal psMovNro As String, pnColPage As Integer, psEmpresaLogo As String, ByVal psEmpresa As String, psEmpresaRuc As String) As String
Dim sTexto As String
Dim lsCadPrint As String
Dim n As Integer
Dim BON As String
Dim BOFF As String
Dim oConect As DConecta
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim lsGlosa As String
Dim lsAgencia As String
Dim lsNomAge As String
Dim CON As String
Dim COFF As String

BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")
Set Rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

Sql = "Select   M.cMovNro,M.cMovDesc, MC.NMOVMONTO AS nMovSaldo , M.cOpeCod, " _
    & "         MD.nDocTpo, D.cDocDesc, MD.cDocNro, " _
    & "         ISNULL(A.cAreaCod,'') + ISNULL(AG.cAgeCod,'') AS cAreaCod, ISNULL(A.cAreaDescripcion,'') + ' ' + ISNULL(AG.cAgeDescripcion,'') AS cAreaDescripcion , MD.dDocFecha," _
    & "         ISNULL(AR.cAreaCod,'') + ISNULL(AGR.cAgeCod,'') AS cAreaCodPers, ISNULL(AR.cAreaDescripcion,'') + ' ' + ISNULL(AGR.cAgeDescripcion,'') AS cAreaDesPers , MD.dDocFecha," _
    & "         C.nConsValor  ,C.cConsDescripcion , P.cPersCod, P.cPersNombre, " _
    & "         ISNULL((Select cPersIDnro FROM PersID  WHERE cPersCod = P.cPersCod and cPersIDTpo =" & gPersIdDNI & "),'') AS DNI , " _
    & "         ISNULL((Select cPersIDnro FROM PersID  WHERE cPersCod = P.cPersCod and cPersIDTpo =" & gPersIdRUC & "),'') AS RUC ,  " _
    & "         MCH.cAreaCod + ' ' + MCH.cAgeCod as cAreaCH, ISNULL(MCH.nProcNro,'') AS nProcCh, A1.cAreaDescripcion + ' ' + ISNULL(AG1.cAgeDescripcion,'')   AS cDescCH " _
    & " From    Mov m JOIN MOVCONT MC ON MC.nMOVNRO=M.nMOVNRO   " _
    & "         Join MovARendir MR  ON MR.nMovNro = M.nMovNro " _
    & "         JOIN MovDoc MD      ON MD.nMovNro = MR.nMovNro " _
    & "         JOIN " & vsBaseComunes & "Documento D    ON D.nDocTpo = MD.nDocTpo " _
    & "         LEFT JOIN " & vsBaseComunes & "Areas A   ON A.cAreaCod = MR.cAreaCod " _
    & "         LEFT JOIN " & vsBaseComunes & "Agencias AG   ON AG.cAgeCod = MR.cAgeCod " _
    & "         JOIN " & vsBasePesonas & "Persona P      ON P.cPersCod = MR.cPersCod " _
    & "         JOIN RRHH RH      ON P.cPersCod = RH.cPersCod " _
    & "         LEFT JOIN " & vsBaseComunes & "Areas AR   ON AR.cAreaCod = RH.cAreaCodActual " _
    & "         LEFT JOIN " & vsBaseComunes & "Agencias AGR   ON AGR.cAgeCod = RH.cAgenciaActual    " _
    & "         JOIN " & vsBaseComunes & "Constante C    ON MR.cTpoARendir = C.nConsValor  AND C.nCONSCOD = '" & ConstanteCabecera.gArendirTipo & "' " _
    & "         LEFT JOIN MOVCAJACHICA MCH ON MCH.nMovNro= M.nMovNro " _
    & "         LEFT JOIN " & vsBaseComunes & "AREAS A1 ON A1.CAREACOD = MCH.cAreaCod " _
    & "         LEFT JOIN " & vsBaseComunes & "Agencias AG1   ON AG1.cAgeCod = MCH.cAgeCod" _
    & " Where M.cMovNro ='" & psMovNro & "' and M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "')"

Set Rs = oConect.CargaRecordSet(Sql)
If Rs.EOF Then Exit Function

If Rs!nConsValor = gArendirTipoCajaChica Then
    sTexto = "RECIBO DE A RENDIR CUENTA DE CAJA CHICA  NRO. " & Rs!cDocNro
Else
    sTexto = "R E C I B O   D E   A   R E N D I R   C U E N T A   NRO. " & Rs!cDocNro
End If
    lsCadPrint = ImpreCabAsiento(pnColPage, Rs!dDocFecha, psEmpresaLogo, Rs!COPECOD, Rs!cMovNro, sTexto, , , , 2)
    ' Definición de Cabecera 2
    If Rs!nConsValor = gArendirTipoCajaChica Then
        lsCadPrint = lsCadPrint + "  CAJA CHICA   : " & BON & ImpreFormat(Rs!cAreaCH & "-" & Rs!nProcCh, 15) & ImpreCarEsp(Rs!cDescCH) & BOFF & oImpresora.gPrnSaltoLinea
    Else
        lsCadPrint = lsCadPrint + "  A Rendir de  : " & BON & ImpreFormat(Rs!cAreaCod, 15) & ImpreCarEsp(Rs!CAREADESCRIPCION) & BOFF & oImpresora.gPrnSaltoLinea
    End If
    lsCadPrint = lsCadPrint + "  Persona      : " & BON & ImpreFormat(Rs!cPersCod, 15) & ImpreCarEsp(PstaNombre(Rs!cPersNombre)) & BOFF & oImpresora.gPrnSaltoLinea
    lsCadPrint = lsCadPrint + "  Area/Agencia : " & BON & ImpreFormat(Rs!cAreaCodPers, 15) & ImpreCarEsp(Rs!cAreaDesPers) & BOFF & oImpresora.gPrnSaltoLinea
    lsCadPrint = lsCadPrint + "  Cargo        : " & oImpresora.gPrnSaltoLinea
    lsCadPrint = lsCadPrint + "  Importe      : " & BON & ConvNumLet(Rs!nMovSaldo, False) & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    If Mid(Rs!COPECOD, 3, 1) = gMonedaExtranjera Then
        'lsCadPrint = lsCadPrint + "                 Tipo de Cambio : " & txtTipCambio & oImpresora.gPrnSaltoLinea
    End If
    lsCadPrint = lsCadPrint + ImpreGlosa(Trim(Rs!cMovDesc), pnColPage, "  Concepto     : ")
    lsCadPrint = lsCadPrint + oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadPrint = lsCadPrint + BON + oImpresora.gPrnSaltoLinea
    lsCadPrint = lsCadPrint + Space(20) & "___________________________" & oImpresora.gPrnSaltoLinea
    lsCadPrint = lsCadPrint + Space(20) & CON & Centra(Mid(PstaNombre(Rs!cPersNombre), 1, 46), 46) & COFF & oImpresora.gPrnSaltoLinea
    lsCadPrint = lsCadPrint + Space(20) & Centra("  L.E. " & Trim(Rs!DNI), 28) & BOFF & oImpresora.gPrnSaltoLinea
    lsCadPrint = lsCadPrint + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
    
    lsGlosa = "De acuerdo a la Política existente, la rendición del presente recibo se deberá hacer en un plazo no mayor de 48 horas, bajo responsabilidad." & oImpresora.gPrnSaltoLinea _
           & "Los Comprobantes de gastos serán a nombre de la " & BON & Trim(psEmpresa) & BOFF & ", RUC " & BON & psEmpresaRuc & BOFF & oImpresora.gPrnSaltoLinea & " "
    lsCadPrint = lsCadPrint + ImpreGlosa(lsGlosa, pnColPage, " NOTA : ") + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
    If Rs!nConsValor <> gArendirTipoCajaChica Then
        lsCadPrint = lsCadPrint + Centra("____________________          _____________________", 80) + oImpresora.gPrnSaltoLinea
        lsCadPrint = lsCadPrint + Centra(" Vo Bo JEFE DE AREA              Vo Bo GERENCIA    ", 80) + oImpresora.gPrnSaltoLinea
    Else
        lsCadPrint = lsCadPrint + Centra("____________________", 80) + oImpresora.gPrnSaltoLinea
        lsCadPrint = lsCadPrint + Centra(" Vo Bo JEFE DE AREA ", 80) + oImpresora.gPrnSaltoLinea
    End If
    lsCadPrint = lsCadPrint + oImpresora.gPrnSaltoLinea
    ImprimeReciboARendir = lsCadPrint
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

Dim oIni As ClasIni
Set oIni = New ClasIni
vsConexion = oIni.CadenaConexion
vsBaseComunes = oIni.BaseComunes
vsBasePesonas = oIni.BasePersonas
'sLpt = "LPT1"
'BON = PrnSet("B+")
'BOFF = PrnSet("B-")
'CON = PrnSet("C+")
'COFF = PrnSet("C-")
Set oIni = Nothing
End Sub
Private Function ImprimeBilletaje(ByVal psMoneda As Moneda, ByVal psMovNro As String, ByVal pnColPage As Integer, Optional ByVal pbCondensado As Boolean = True) As String
Dim sSql As String
Dim Rs   As New ADODB.Recordset
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
Set Rs = oEfec.GetEfectivoMovImp(psMovNro, True)
If Not Rs.EOF And Not Rs.BOF Then
    If Rs!Moneda = gMonedaExtranjera Then
        lsSimbolo = "$."
    Else
        '''lsSimbolo = "S/." 'MARG ERS044-2016
        lsSimbolo = gcPEN_SIMBOLO
    End If
    sTexto = sTexto + BON & PrnSet("N+") & "DESCOMPOSICION DE EFECTIVO : INGRESO" & BOFF & PrnSet("N-") & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + IIf(pbCondensado, CON, "") + String(pnColPage - 3, "=") & oImpresora.gPrnSaltoLinea
    nMontoM = 0: nMontoB = 0
    nMontoTot = 0
    sTexto = sTexto + BON & ImpreFormat(" BILLETES ", 22, 7) & ImpreFormat("MONTO", 15) _
         & ImpreFormat(" MONEDAS ", 14, 5) & ImpreFormat("MONTO", 10) + BOFF & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + String(pnColPage - 3, "-") + BOFF & oImpresora.gPrnSaltoLinea
    ReDim otBilletajes(0)
    lnFilaBill = 0
    lnFilaMon = 0
    Do While Not Rs.EOF
        If Rs!billetes <> 0 Then
            lnFilaBill = lnFilaBill + 1
            If lnFilaBill > lnFilaMon Then
            ReDim Preserve otBilletajes(lnFilaBill)
            End If
            otBilletajes(lnFilaBill).lsBilletes = Rs!billetes
            otBilletajes(lnFilaBill).lnMontoBil = Rs!MontoBilletes
        End If
        If Rs!Monedas <> 0 Then
            lnFilaMon = lnFilaMon + 1
            If lnFilaMon > lnFilaBill Then
            ReDim Preserve otBilletajes(lnFilaMon)
            End If
            otBilletajes(lnFilaMon).lsMonedas = Rs!Monedas
            otBilletajes(lnFilaMon).lnMontoMon = Rs!MontoMonedas
        End If
        nMontoM = nMontoM + Rs!MontoMonedas
        nMontoB = nMontoB + Rs!MontoBilletes
        nMontoTot = nMontoTot + Rs!MontoMonedas + Rs!MontoBilletes
       Rs.MoveNext
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
        sTexto = sTexto + lsBilletes + lsMonedas & oImpresora.gPrnSaltoLinea
    Next
    sTexto = sTexto + String(pnColPage - 3, "-") & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + BON & ImpreFormat("SUBTOTAL :", 12, 3) & ImpreFormat(lsSimbolo, 3, 1) & ImpreFormat(nMontoB, 15, 2, True) _
         & ImpreFormat("SUBTOTAL", 10, 8) & ImpreFormat(lsSimbolo, 3, 1) & ImpreFormat(nMontoM, 10, 2, True) + BOFF & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + String(pnColPage - 3, "=") & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + Space(45) & BON & "TOTAL :  " _
           & lsSimbolo & " " & PrnVal(nMontoTot, 14, 2) & BOFF & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + Space(45) & String(31, "=") + COFF & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + "" & oImpresora.gPrnSaltoLinea
End If
Rs.Close
Set Rs = Nothing

Set Rs = oEfec.GetEfectivoMovImp(psMovNro, False)
If Not Rs.EOF And Not Rs.BOF Then
    If Rs!Moneda = gMonedaExtranjera Then
        lsSimbolo = "$."
    Else
        '''lsSimbolo = "S/." 'MARG ERS044-2016
        lsSimbolo = gcPEN_SIMBOLO 'MARG ERS044-2016
    End If
    sTexto = sTexto + BON & PrnSet("N+") & "DESCOMPOSICION DE EFECTIVO : SALIDA" & BOFF & PrnSet("N-") & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + IIf(pbCondensado, CON, "") + String(pnColPage - 3, "=") & oImpresora.gPrnSaltoLinea
    nMontoM = 0: nMontoB = 0
    nMontoTot = 0
    sTexto = sTexto + BON & ImpreFormat(" BILLETES ", 22, 7) & ImpreFormat("MONTO", 15) _
         & ImpreFormat(" MONEDAS ", 14, 5) & ImpreFormat("MONTO", 10) + BOFF & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + String(pnColPage - 3, "-") + BOFF & oImpresora.gPrnSaltoLinea
    ReDim otBilletajes(0)
    lnFilaBill = 0
    lnFilaMon = 0
    Do While Not Rs.EOF
        If Rs!billetes <> 0 Then
            lnFilaBill = lnFilaBill + 1
            If lnFilaBill > lnFilaMon Then
            ReDim Preserve otBilletajes(lnFilaBill)
            End If
            otBilletajes(lnFilaBill).lsBilletes = Rs!billetes
            otBilletajes(lnFilaBill).lnMontoBil = Rs!MontoBilletes
        End If
        If Rs!Monedas <> 0 Then
            lnFilaMon = lnFilaMon + 1
            If lnFilaMon > lnFilaBill Then
            ReDim Preserve otBilletajes(lnFilaMon)
            End If
            otBilletajes(lnFilaMon).lsMonedas = Rs!Monedas
            otBilletajes(lnFilaMon).lnMontoMon = Rs!MontoMonedas
        End If
        nMontoM = nMontoM + Rs!MontoMonedas
        nMontoB = nMontoB + Rs!MontoBilletes
        nMontoTot = nMontoTot + Rs!MontoMonedas + Rs!MontoBilletes
       Rs.MoveNext
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
        sTexto = sTexto + lsBilletes + lsMonedas & oImpresora.gPrnSaltoLinea
    Next
    sTexto = sTexto + String(pnColPage - 3, "-") & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + BON & ImpreFormat("SUBTOTAL :", 12, 3) & ImpreFormat(lsSimbolo, 3, 1) & ImpreFormat(nMontoB, 15, 2, True) _
         & ImpreFormat("SUBTOTAL", 10, 8) & ImpreFormat(lsSimbolo, 3, 1) & ImpreFormat(nMontoM, 10, 2, True) + BOFF & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + String(pnColPage - 3, "=") & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + Space(45) & BON & "TOTAL :  " _
           & lsSimbolo & " " & PrnVal(nMontoTot, 14, 2) & BOFF & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + Space(45) & String(31, "=") + COFF & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + "" & oImpresora.gPrnSaltoLinea
End If
Rs.Close
Set Rs = Nothing
ImprimeBilletaje = sTexto
Exit Function
ErrBillete:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeBilletaje Method")
End Function

Public Function ImprimeOrdenPago(sTexto As String) As Boolean
Dim I As Integer
   ImprimeOrdenPago = True
   lbCancela = True
   
   If lbCancela = False Then
      For I = 1 To lnNumCopias
         ImpreBegin False, 1
         Print #ArcSal, sTexto;
         Print #ArcSal, oImpresora.gPrnCondensadaOFF;   'Retorna al tipo de letra normal
         Close ArcSal
      Next I
   Else
      ImprimeOrdenPago = False
   End If
End Function

Public Function ImprimeReciboIngresoEgreso(ByVal psMovNro As String, pdFecha As Date, ByVal psGlosa As String, _
                                            ByVal psEmpresaLogo As String, ByVal psOpeCod As String, _
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
Dim n As Integer
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim oConect As DConecta
Dim psAgeCod  As String, psAgeDesc As String
Dim psAreaCod  As String, psAreaDesc As String
Dim psPersNombre As String
    

Set oConect = New DConecta
Set Rs = New ADODB.Recordset

If oConect.AbreConexion = False Then Exit Function

Sql = " SELECT  P.cPersCod, P.cPersNombre , " _
    & "         ISNULL(A.cAreaCod,'') cAreaCod,  ISNULL(A.cAreaDescripcion,'') cAreaDescripcion, " _
    & "         ISNULL(AG.cAgeCod,'') cAgeCod, ISNULL(AG.cAgeDescripcion,'') cAgeDescripcion  " _
    & " FROM    PERSONA P " _
    & "         JOIN RRHH   RH ON RH.CPERSCOD = P.CPERSCOD " _
    & "         LEFT JOIN AREAS  A ON A.CAREACOD = RH.CAREACOD " _
    & "         LEFT JOIN AGENCIAS AG ON AG.CAGECOD = RH.cAgenciaActual " _
    & " WHERE   P.CPERSCOD ='" & psPersCod & "'"

Set Rs = oConect.CargaRecordSet(Sql)
If Not Rs.EOF And Not Rs.BOF Then
    psAgeCod = Rs!cAgeCod
    psAgeDesc = Rs!cAgeDescripcion
    psAreaCod = Rs!cAreaCod
    psAreaDesc = Rs!CAREADESCRIPCION
    psPersNombre = PstaNombre(Rs!cPersNombre)
Else
    Exit Function
End If
Rs.Close
Set Rs = Nothing

BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")

On Error GoTo ImprimeReciboIngresoEgresoErr
    nAncho = pnColPage
    sTexto = ""
    If pbHabilitacionCH = False Then
        If pbIngreso Then
            sTexto = sTexto + ImpreCabAsiento(pnColPage, pdFecha, psEmpresaLogo, psOpeCod, psMovNro, " R E C I B O   D E   I N G R E S O ", , , , 1) & oImpresora.gPrnSaltoLinea
        Else
            sTexto = sTexto + ImpreCabAsiento(pnColPage, pdFecha, psEmpresaLogo, psOpeCod, psMovNro, " R E C I B O   D E   E G R E S O ", , , , 1) & oImpresora.gPrnSaltoLinea
        End If
    Else
        sTexto = sTexto + ImpreCabAsiento(pnColPage, pdFecha, psEmpresaLogo, psOpeCod, psMovNro, "H A B I L I T A C I O N    D E    C A J A    C H I C A", , , , 1) & oImpresora.gPrnSaltoLinea
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
    sTexto = sTexto + "  Importe      : " & BON & ConvNumLet(lnMonto, False) & BOFF & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + ImpreGlosa(psGlosa, pnColPage, "  Concepto     : ") & oImpresora.gPrnSaltoLinea
    'sTexto = sTexto + "" & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea
    'sTexto = sTexto + BON & Space(20) & "_________________________" & oImpresora.gPrnSaltoLinea
    'Texto = sTexto + Space(20) & Centra("Vo Bo. CAJA ", 25) & BOFF & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea
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
                                         ByVal pdFecSis As Date, ByVal psCtaContArendir As String, ByVal psCtaContPendiente As String, ByVal psSimbolo As String, _
                                         ByVal pnMoneda As Moneda, ByVal pnTipoArendir As ArendirTipo) As String
Dim n As Integer
Dim P As Integer
Dim nLin As Integer
Dim Cont As Integer
Dim lsTexto As String
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim oNARendir As NARendir
'Dim BON As String
'Dim BOFF As String
'Dim CON As String
'Dim COFF As String
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
Set Rs = New ADODB.Recordset
Set Rs = oNARendir.GetDatosArendir(psMovNroSol, psCtaContArendir, psCtaContPendiente, pnMoneda, pnTipoArendir)
If Not Rs.EOF And Not Rs.BOF Then
    lnSaldo = Rs!nMovSaldo
    lnSust = 0
    'Identificamos los Documentos Sustentatorios del Recibo de A rendir
    Do While Not Rs.EOF
        lsTexto = lsTexto + CabeceraPlanillaRendicion(pnTipoArendir, " DECLARACION DE A RENDIR CUENTA ", nLin, P, pnLinPage, pnColPage, psNomCmac, pdFecSis, Rs!cDocNro, IIf(IsNull(Rs!cDocDescAtenc), "ATENCION", Rs!cDocDescAtenc), _
                IIf(IsNull(Rs!cDocNroAtenc), "EFECTIVO", Rs!cDocNroAtenc), IIf(IsNull(Rs!dFechaDocAtenc), Mid(Rs!CMOVATENC, 7, 2) & "/" & Mid(Rs!CMOVATENC, 5, 2) & "/" & Mid(Rs!CMOVATENC, 1, 4), Rs!dFechaDocAtenc), psSimbolo, Rs!ImporteARendir, Rs!cAgeCod, Rs!cAgeDescripcion, _
                Rs!cAreaCod, Rs!CAREADESCRIPCION, Rs!cPersCod, Rs!cPersNombre, Rs!cAreaCodArendir, Rs!cAgeCodArendir, Rs!cAreaArendir, Rs!cAgeArendir)

        lsTexto = lsTexto + CON & "   " & Format(Rs.Bookmark, "000") & "  " & Rs!dDocFechaSust & "  " & _
                 Left(Rs!cDocNroAbrevSust & Space(4), 4) & " " & Left(Rs!cDocNroSust & Space(14), 14) & " " & _
                 Left(PstaNombre(Rs!cPersCodSust) & Space(40), 40) & _
                 "  " & Left(Rs!cDescSust & Space(35), 35) & " " & Right(Space(14) & Format(Rs!nDocImporte, "#,##0.00"), 14) & COFF & oImpresora.gPrnSaltoLinea
                 
        nLin = nLin + 1
        lnSust = lnSust + Rs!nDocImporte
        Rs.MoveNext
    Loop
    lsTexto = lsTexto + CON
    lsTexto = lsTexto + " ------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + "                                                                              TOTAL SUSTENTADO                 :  " & psSimbolo & " " & PrnVal(CCur(Format(lnSust, "##,###0.00")), 14, 2) & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + "" & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + BON + "                                                                              SALDO POR RENDIR EN CAJA         :  " & psSimbolo & " " & PrnVal(CCur(Format(lnSaldo, "##,###0.00")), 14, 2) + BOFF + oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + " ====================================================================================================================================" & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + COFF + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + Centra("______________________    ______________________   ______________________", pnColPage) & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + Centra("   Vo Bo  USUARIO           Vo Bo JEFE DE AREA         Vo Bo GERENCIA    ", pnColPage) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
End If
Rs.Close
Set Rs = Nothing
ImprimePlanillaRendicion = lsTexto
Set oNARendir = Nothing
End Function
Private Function CabeceraPlanillaRendicion(ByVal pnTipoArendir As ArendirTipo, ByVal psTitulo As String, ByRef nLin As Integer, ByRef P As Integer, ByVal pnLinPage As Integer, _
                                           ByVal pnColPage As Integer, ByVal psNomCmac As String, _
                                           ByVal pdFecSis As Date, ByVal psNroDocArendir As String, ByVal psDescDocAtend As String, _
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
   lsTexto = lsTexto + ImpreFormat(psNomCmac, 60) & pdFecSis & " - " & Format(Time, "hh:mm:ss") + oImpresora.gPrnSaltoLinea
   lsTexto = lsTexto + " CONTABILIDAD    " & Space(55) & "Pag. " & Format(P, "000") + oImpresora.gPrnSaltoLinea
   lsTexto = lsTexto + BON & Centra(psTitulo, pnColPage) & BOFF + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
   If pnTipoArendir = gArendirTipoCajaGeneral Then
        lsTexto = lsTexto + CON & "  Recibo de A rendir: " & BON & psNroDocArendir & BOFF + oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + "  A Rendir de.......: " & BON & psAreaCodARendir + psAgeCod & " " & psAreaDescArendir + " " + psAgeDescArendir & BOFF + oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + "  " & Mid(psDescDocAtend & Space(18), 1, 18) & ": " & BON & psNroDocAtend & BOFF & "       Fecha : " & BON & psDocAtencFecha & BOFF & _
            Space(15) & "Monto : " & BON & Right(Space(18) & psSimbolo & " " & Format(ImporteARendir, "#,#0.00"), 18) & BOFF + oImpresora.gPrnSaltoLinea
   Else
        lsTexto = lsTexto + CON & "  Recibo de A rendir: " & BON & psNroDocArendir & BOFF + Space(15) & "Monto : " & BON & Right(Space(18) & psSimbolo & " " & Format(ImporteARendir, "#,#0.00"), 18) & BOFF + oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + "  A Rendir de.......: " & BON & psAreaCodARendir + psAgeCodArendir & " " & psAreaDescArendir + " " + psAgeDescArendir & BOFF + oImpresora.gPrnSaltoLinea
   End If
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
        '''lsTexto = lsTexto + COFF + BON & "ASIENTO CONTABLE EN " & IIf(psSimbolo = "S/.", "M.N. :", "M.E. :") & BOFF & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
        lsTexto = lsTexto + COFF + BON & "ASIENTO CONTABLE EN " & IIf(psSimbolo = gcPEN_SIMBOLO, "M.N. :", "M.E. :") & BOFF & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
        lsTexto = lsTexto + String(pnColPage - 1, "=") & oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + "  " & BON & "Cuenta Contable" & Space(43) & "DEBE       HABER" & BOFF & oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + String(pnColPage - 1, "-") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    End If
    nLin = 15
End If
CabeceraPlanillaRendicion = lsTexto
End Function
Public Function ImprimeAsientoSustArendir(ByVal psMovNroAtenc As String, ByVal pnColPage As Integer, ByVal pnLinPage As Integer, ByVal psNomCmac As String, _
                                         ByVal pdFecSis As Date, ByVal psCtaContArendir As String, ByVal psSimbolo As String, ByVal pnTipoArendir As ArendirTipo)
Dim lsTexto As String
Dim Rs As ADODB.Recordset
Dim oArendir As NARendir

Dim oGen As DGeneral
Dim BON As String: Dim BOFF As String
Dim CON As String:   Dim COFF As String

Dim lnDebe As Currency: Dim lnHaber As Currency
Dim nLin As Integer
Dim P As Integer


Set oGen = New DGeneral

Set oArendir = New NARendir
Set Rs = New ADODB.Recordset

BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")
nLin = 66
P = 0
Set Rs = oArendir.GetAsientosSustARendir(psMovNroAtenc, psCtaContArendir, pnTipoArendir)
If Not Rs.EOF And Not Rs.BOF Then
    Do While Not Rs.EOF
        
        lsTexto = lsTexto + CabeceraPlanillaRendicion(pnTipoArendir, "SUSTENTACION DE A RENDIR CUENTA ", nLin, P, pnLinPage, pnColPage, psNomCmac, pdFecSis, Rs!cDocNro, IIf(Rs!cDocDescAtenc = "", "ATENCION", Rs!cDocDescAtenc), _
               IIf(Rs!cDocNroAtenc = "", "EFECTIVO", Rs!cDocNroAtenc), IIf(IsNull(Rs!dFechaDocAtenc), Mid(Rs!CMOVATENC, 7, 2) & "/" & Mid(Rs!CMOVATENC, 5, 2) & "/" & Mid(Rs!CMOVATENC, 1, 4), Rs!dFechaDocAtenc), psSimbolo, Rs!ImporteARendir, Rs!cAgeCod, Rs!cAgeDescripcion, _
               Rs!cAreaCod, Rs!CAREADESCRIPCION, Rs!cPersCod, Rs!cPersNombre, Rs!cAreaCodArendir, Rs!cAgeCodArendir, Rs!cAreaArendir, Rs!cAgeArendir, False)
               
        lsTexto = lsTexto + CON + "   " & Justifica(Rs!CtaContSust, 18) & "  " & Justifica(oGen.CuentaNombre(Rs!CtaContSust), 66) & " " & psSimbolo & " " & PrnVal(Rs!Debe, 14, 2) & " " & psSimbolo & " " & PrnVal(Rs!Haber, 14, 2) & oImpresora.gPrnSaltoLinea
        lnDebe = lnDebe + Rs!Debe
        lnHaber = lnHaber + Rs!Haber
        nLin = nLin + 1
        Rs.MoveNext
    Loop
    lsTexto = lsTexto + COFF + oImpresora.gPrnSaltoLinea + String(pnColPage - 1, "-") & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + CON
    lsTexto = lsTexto + Space(75) & "TOTALES        " & psSimbolo & Right(Space(15) & Format(lnDebe, "###,###,##0.00"), 15) & "  " & psSimbolo & Right(Space(15) & Format(lnHaber, "###,###,##0.00"), 15) & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + COFF
    lsTexto = lsTexto + String(pnColPage - 1, "=") & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto + ImprePiePag(pnColPage, "89", "VoBo  SOLICITANTE")
End If
Rs.Close
Set Rs = Nothing
Set oGen = Nothing
Set oArendir = Nothing
ImprimeAsientoSustArendir = lsTexto

End Function
Public Function ImprimeDetDocCtaCont(ByVal psMovNroSol As String, ByVal pnColPage As Integer, ByVal pnLinPage As Integer, _
                                  ByVal psCtaContArendir As String, ByVal psCtaContPendiente As String, ByVal psSimbolo As String, pnTipoArendir As ArendirTipo)

Dim lsTexto As String
Dim Rs As ADODB.Recordset
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
Set Rs = New ADODB.Recordset

BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")

lsTexto = ""

Set Rs = oArendir.GetDetDocCtaCont(psMovNroSol, psCtaContArendir, psCtaContPendiente, pnTipoArendir)

If Not Rs.EOF And Not Rs.BOF Then
    sCabe = CON + String(125, "=") & oImpresora.gPrnSaltoLinea _
      & "  D O C U M E N T O               C O N C E P T O                                                DEBE           HABER" & oImpresora.gPrnSaltoLinea _
      & "TIPO  NUMERO        FECHA " & oImpresora.gPrnSaltoLinea _
      & String(125, "-") + COFF & oImpresora.gPrnSaltoLinea

    nLi = pnLinPage: P = 0
    lsTexto = PrnSet("MI", 4)
    If nLi > pnLinPage - 6 Then
        lsTexto = lsTexto + Cabecera("DETALLE DE DOCUMENTOS POR CUENTA CONTABLE", P, psSimbolo, pnColPage, sCabe) & oImpresora.gPrnSaltoLinea
        nLi = 9
    End If
    Do While Not Rs.EOF
        lsTexto = lsTexto + CON + "CUENTA CONTABLE: " & Justifica(Rs!CtaContSust, 18) & " " & Justifica(oGen.CuentaNombre(Rs!CtaContSust), 80) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        nLi = nLi + 2
        lsCtaCont = Rs!CtaContSust
        lnDebe = 0: lnHaber = 0
        Do While lsCtaCont = Rs!CtaContSust
            lnDebe = lnDebe + Rs!Debe
            lnHaber = lnHaber + Rs!Haber
            lsTexto = lsTexto + Justifica(Rs!cDocAbrev, 3) & " " & Justifica(Rs!cDocNro, 16) & " " & Justifica(Rs!dDocFecha, 10) & "  " & Justifica(Rs!cMovDesc, 60) & " " & PrnVal(Rs!Debe, 14, 2) & " " & PrnVal(Rs!Haber, 14, 2) & oImpresora.gPrnSaltoLinea
            nLi = nLi + 1
            If nLi > pnLinPage - 6 Then
                lsTexto = lsTexto + Cabecera("DETALLE DE DOCUMENTOS POR CUENTA CONTABLE", P, psSimbolo, pnColPage, sCabe) & oImpresora.gPrnSaltoLinea
                nLi = 9
            End If
            Rs.MoveNext
            If Rs.EOF Then Exit Do
        Loop
        lsTexto = lsTexto + String(125, "-") & oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + Space(94) & PrnVal(lnDebe, 14, 2) & " " & PrnVal(lnHaber, 14, 2) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        nLi = nLi + 3
    Loop
End If
Rs.Close
Set Rs = Nothing
Set oGen = Nothing
Set oArendir = Nothing

ImprimeDetDocCtaCont = lsTexto
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
Dim lnCol As Long
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
lnCol = 78

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
lsTexto = CON + BON + PrnSet("I+") + ImpreFormat(lsTitulo, lnCol, 50) & ImpreFormat(lsTitulo, lnCol, 10) + PrnSet("I-") + BOFF + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
lsTexto = lsTexto + ImpreFormat(Trim(psAgeDesc), lnCol, 3) & ImpreFormat(Trim(psAgeDesc), lnCol, lnColIzq) + oImpresora.gPrnSaltoLinea
lnLineas = lnLineas + 1
lsTexto = lsTexto + ImpreFormat(CentrarCadena(String(33, "=") + " Concepto " + String(35, "="), lnCol), lnCol, 3) + ImpreFormat(CentrarCadena(String(33, "=") + " Concepto " + String(35, "="), lnCol), lnCol, lnColIzq) & oImpresora.gPrnSaltoLinea
lnLineas = lnLineas + 1
'Impresion de la glosa
lsGlosa = JustificaTexto(psGlosa, lnCol - 6)
lsGlosaLin = "": lnPos = 0: Char = ""
lnLineasGlosa = 0
For I = 1 To Len(lsGlosa)
    Char = Mid(lsGlosa, I, 1)
    lsGlosaLin = lsGlosaLin & Char
    If Char = oImpresora.gPrnSaltoLinea Or I = Len(lsGlosa) Then
        lsGlosaLin = Mid(lsGlosaLin, 1, IIf(I = Len(lsGlosa), Len(lsGlosaLin), Len(lsGlosaLin) - 1))
        lsTextoGlosa = lsTextoGlosa & ImpreFormat(Trim(lsGlosaLin), lnCol, lnColIzq - 7) & ImpreFormat(Trim(lsGlosaLin), lnCol, lnColIzq) & oImpresora.gPrnSaltoLinea
        lnLineas = lnLineas + 1
        lnLineasGlosa = lnLineasGlosa + 1
        If lnLineasGlosa = 3 Then
            Exit For
        End If
        lsGlosaLin = ""
    End If
Next
lsTexto = lsTexto + lsTextoGlosa
lsTexto = lsTexto + ImpreFormat(String(50, "_") & " CARGOS " & String(10, "_") & " ABONOS " & String(14, "_"), lnCol, 3) + ImpreFormat(String(50, "_") & " CARGOS " & String(10, "_") & " ABONOS " & String(14, "_"), lnCol, lnColIzq) & oImpresora.gPrnSaltoLinea
If lnImporteCargo > 0 Then
    lsTexto = lsTexto + ImpreFormat(ImpreFormat("IMPORTE :", 40) & ImpreFormat(lnImporteCargo, 14, 2, True) & ImpreFormat(0, 15, 2, True), lnCol, 3) + ImpreFormat(ImpreFormat("IMPORTE :", 40) & ImpreFormat(lnImporteCargo, 14, 2, True) & ImpreFormat(0, 15, 2, True), lnCol, lnColIzq) & oImpresora.gPrnSaltoLinea
Else
    lsTexto = lsTexto + ImpreFormat(ImpreFormat("IMPORTE :", 40) & ImpreFormat(0, 14, 2, True) & ImpreFormat(lnImporteAbono, 15, 2, True), lnCol, 3) + ImpreFormat(ImpreFormat("IMPORTE :", 40) & ImpreFormat(0, 14, 2, True) & ImpreFormat(lnImporteAbono, 15, 2, True), lnCol, lnColIzq) & oImpresora.gPrnSaltoLinea
End If
lsTexto = lsTexto + ImpreFormat(String(lnCol, "_"), lnCol, 3) + ImpreFormat(String(lnCol, "_"), lnCol, lnColIzq) & oImpresora.gPrnSaltoLinea
lsTexto = lsTexto + ImpreFormat("Sr(es) :", lnCol, 3) + ImpreFormat("Sr(es) :", lnCol, lnColIzq) & oImpresora.gPrnSaltoLinea
lsTexto = lsTexto + BON + ImpreFormat(PstaNombre(psPersNombre), lnCol, lnColIzq + 5) + ImpreFormat(PstaNombre(psPersNombre), lnCol, lnColIzq + 5) & oImpresora.gPrnSaltoLinea
lsTexto = lsTexto + ImpreFormat(psPersDireccion, lnCol, lnColIzq + 5) + ImpreFormat(psPersDireccion, lnCol, lnColIzq + 5) & oImpresora.gPrnSaltoLinea
lsTexto = lsTexto + ImpreFormat(psPersUbiGeo, lnCol, lnColIzq + 5) + ImpreFormat(psPersUbiGeo, lnCol, lnColIzq + 5) + BOFF + oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea

lsTexto = lsTexto + ImpreFormat(String(3, "-") & "FECHA" & String(13, "-") & " MONEDA " & String(5, "-") & " CUENTA " & String(15, "-") + BON + lsTextoTotal + BOFF & String(15, "-"), lnCol + 4, 3) _
            + ImpreFormat(String(3, "-") & "FECHA" & String(13, "-") & " MONEDA " & String(5, "-") & " CUENTA " & String(15, "-") + BON + lsTextoTotal + BOFF & String(8, "-"), lnCol + 4, lnColIzq) & oImpresora.gPrnSaltoLinea

'''lsTexto = lsTexto + ImpreFormat(String(3, " ") & Format(pdFecha & " " & Time, gsFormatoFechaHora) & String(3, " ") & IIf(psMoneda = gMonedaNacional, "SOLES", "DOLARES") & String(3, " ") & psNroCtaAhorros + BON + ImpreFormat(pnImporte, 15, 2, True) + BOFF + ImpreFormat(psCodUser, 5, 4), lnCol + 4, 3) _
'''            + ImpreFormat(String(3, " ") & Format(pdFecha & " " & Time, gsFormatoFechaHora) & String(3, " ") & IIf(psMoneda = gMonedaNacional, "SOLES", "DOLARES") & String(3, " ") & psNroCtaAhorros + BON + ImpreFormat(pnImporte, 15, 2, True) + BOFF + ImpreFormat(psCodUser, 5, 4), lnCol + 4, lnColIzq) & oImpresora.gPrnSaltoLinea 'MARG ERS0442016

lsTexto = lsTexto + ImpreFormat(String(3, " ") & Format(pdFecha & " " & Time, gsFormatoFechaHora) & String(3, " ") & IIf(psMoneda = gMonedaNacional, StrConv(gcPEN_PLURAL, vbUpperCase), "DOLARES") & String(3, " ") & psNroCtaAhorros + BON + ImpreFormat(pnImporte, 15, 2, True) + BOFF + ImpreFormat(psCodUser, 5, 4), lnCol + 4, 3) _
            + ImpreFormat(String(3, " ") & Format(pdFecha & " " & Time, gsFormatoFechaHora) & String(3, " ") & IIf(psMoneda = gMonedaNacional, StrConv(gcPEN_PLURAL, vbUpperCase), "DOLARES") & String(3, " ") & psNroCtaAhorros + BON + ImpreFormat(pnImporte, 15, 2, True) + BOFF + ImpreFormat(psCodUser, 5, 4), lnCol + 4, lnColIzq) & oImpresora.gPrnSaltoLinea

lsTexto = lsTexto + ImpreFormat(String(lnCol, "-"), lnCol, 3) + ImpreFormat(String(lnCol, "-"), lnCol, lnColIzq) + COFF + oImpresora.gPrnSaltoLinea

'ImprimeNotaCargoAbono = lsTexto & oImpresora.gPrnSaltoPagina  & lsTexto
ImprimeNotaCargoAbono = lsTexto
End Function


'---- NUEVAS FUNCIONES SIMON

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
                                       ByVal pnColPage As Integer, ByVal pdFecSis As Date, ByVal psOpeCod As String, _
                                       ByVal psEmpresaLogo As String, ByVal psSimbolo As String, _
                                       ByVal psNroDocViaticos As String, ByVal psNroDocRef As String, ByVal pdFechaDocRef As String, _
                                       ByVal psAreaCod As String, ByVal psAreaDesc As String, _
                                       ByVal psAgeCod As String, ByVal psAgeDesc As String, _
                                       ByVal psPersCod As String, ByVal psPersNombre As String, ByVal psDocPers As String, _
                                       ByVal psPersCargo As String, ByVal psCategoria As String, _
                                       ByVal pbSeleccion As Boolean, ByVal psMovNro As String, _
                                       ByVal psEmpresa As String, ByVal psEmpresaRuc As String) As String
Dim sTexto As String
Dim nAncho As Integer, n As Integer
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
        sTexto = sTexto + PrnSet("MI", 5) & ImpreCabAsiento(pnColPage, pdFecSis, psEmpresaLogo, psOpeCod, lsMovNro, lsTitulo) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Documento Refe.  : " & BON & Mid(psNroDocRef, 1, 50) & BOFF & Space(10) & "Fecha : " & BON & pdFechaDocRef & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Area             : " & BON & psAreaCod & "  " & ImpreCarEsp(psAreaDesc) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Agencia          : " & BON & psAgeCod & "  " & ImpreCarEsp(psAgeDesc) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Persona          : " & BON & psPersCod & "  " & ImpreCarEsp(psPersNombre) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Cargo            : " & BON & Mid(psPersCargo, 1, 50) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Categoria        : " & BON & Mid(psCategoria, 1, 50) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Destino          : " & BON & Justifica(Mid(!Destino, 1, 50), 50) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Transporte Ida   : " & BON & Justifica(Mid(!Ida, 1, 50), 50) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Transporte Vuelta: " & BON & Justifica(Mid(!Vuelta, 1, 50), 50) & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Partida      : " & BON & !Partida & BOFF & Space(10) & "Nro. Días    : " & BON & PrnVal(!Dias, 5, 0) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Importe      : " & BON & ConvNumLet(PrnVal(!Importe, 10, 2), False) & BOFF & oImpresora.gPrnSaltoLinea
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
        
        lsGlosa = "De acuerdo a la Política existente, la rendición del presente recibo se deberá realizar en el plazo establecido, bajo responsabilidad." & oImpresora.gPrnSaltoLinea _
               & "Los Comprobantes de gastos serán a nombre de " & BON & Trim(psEmpresa) & BOFF & ", RUC " & BON & Trim(psEmpresaRuc) & BOFF & oImpresora.gPrnSaltoLinea & " "
        sTexto = sTexto + ImpreGlosa(lsGlosa, pnColPage, " NOTA : ") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + Centra("____________________          _____________________", 80) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + Centra("    Vo Bo RRHH                    Vo Bo GERENCIA    ", 80) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & PrnSet("MI", 0)
        If .EOF Then
            Exit Do
        End If
        rsAux.MoveNext
        .MoveNext
    Loop
    ImprimirReciboViatico = sTexto
End With
   
End Function
Public Function ImprimirReciboViaticoData(ByVal pnColPage As Integer, ByVal pdFecSis As Date, ByVal psOpeCod As String, _
                                       ByVal psEmpresaLogo As String, ByVal psSimbolo As String, _
                                       ByVal psEmpresa As String, ByVal psEmpresaRuc As String, _
                                       ByVal psMovNroViat As String) As String
Dim sTexto As String
Dim nAncho As Integer, n As Integer
Dim lsGlosa As String
Dim I As Integer
Dim lnImporte As Currency
Dim lsMovNro As String
Dim lnTotal As Currency
Dim Rs As ADODB.Recordset
Dim rsC As ADODB.Recordset
Dim oArendir As NARendir
Dim lsTitulo As String

Set oArendir = New NARendir
Set Rs = New ADODB.Recordset
Set rsC = New ADODB.Recordset

sTexto = ""
nAncho = pnColPage
Set Rs = oArendir.GetDatosTotalesViaticos(psMovNroViat)
sTexto = ""
With Rs
    Do While Not .EOF
        lsMovNro = Rs!cMovNro
        If sTexto <> "" Then
            sTexto = sTexto + oImpresora.gPrnSaltoPagina
        End If
        lsTitulo = "RECIBO DE VIATICOS A RENDIR CUENTA NRO. " & !DOCVIAT
        sTexto = sTexto + PrnSet("MI", 5) & ImpreCabAsiento(pnColPage, pdFecSis, psEmpresaLogo, !COPECOD, lsMovNro, lsTitulo) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Documento Refe.  : " & BON & Mid(!NRODOCREF, 1, 50) & BOFF & Space(10) & "Fecha : " & BON & !dDocFecha & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  A Rendir de...   : " & BON & !cCodAreaArendir & !cCodAgeArendir & "  " & ImpreCarEsp(!cAreaArendir) & " " & ImpreCarEsp(!cAgeArendir) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Persona          : " & BON & !cPersCod & "  " & ImpreCarEsp(PstaNombre(!cPersNombre)) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Area             : " & BON & !cAreaCod & "  " & ImpreCarEsp(!CAREADESCRIPCION) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Agencia          : " & BON & !cAgeCod & "  " & ImpreCarEsp(!cAgeDescripcion) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Cargo            : " & BON & Mid(!cRHCargoDescripcion, 1, 50) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Categoria        : " & BON & Mid(!Categoria, 1, 50) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Destino          : " & BON & Justifica(Mid(!Destino, 1, 50), 50) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Transporte Ida   : " & BON & Justifica(Mid(!TransIda, 1, 50), 50) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Transporte Vuelta: " & BON & Justifica(Mid(!TrasnVuelta, 1, 50), 50) & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Partida      : " & BON & !dPartida & BOFF & Space(10) & "Nro. Días    : " & BON & PrnVal(!nMovViaticosDias, 5, 0) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "  Importe      : " & BON & ConvNumLet(PrnVal(!nMovMonto, 10, 2), False) & BOFF & oImpresora.gPrnSaltoLinea
        lsGlosa = !Motivo
        sTexto = sTexto + ImpreGlosa(lsGlosa, pnColPage, "  Concepto     : ") & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "   ==================================================================" & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "     C O N C E P T O                                   IMPORTE (" & psSimbolo & ")" & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + "   ------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
        lnTotal = 0
        Set rsC = oArendir.GetDetalleCostosViaticos(Rs!cMovNro)
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
        sTexto = sTexto + Space(20) & BON & Centra(Mid(PstaNombre(!cPersNombre), 1, 49), 50) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + Space(20) & Centra("  LE/DNI " & !DNI, 50) & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        
        lsGlosa = "De acuerdo a la Política existente, la rendición del presente recibo se deberá realizar en el plazo establecido, bajo responsabilidad." & oImpresora.gPrnSaltoLinea _
               & "Los Comprobantes de gastos serán a nombre de " & BON & Trim(psEmpresa) & BOFF & ", RUC " & BON & Trim(psEmpresaRuc) & BOFF & oImpresora.gPrnSaltoLinea & " "
        sTexto = sTexto + ImpreGlosa(lsGlosa, pnColPage, " NOTA : ") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + Centra("____________________          _____________________", 80) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + Centra("    Vo Bo RRHH                    Vo Bo GERENCIA    ", 80) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto + oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & PrnSet("MI", 0)
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
Public Function GetImprimeDocRendicionCH(ByVal pnLinPage As Integer, ByVal psNomCmact As String, ByVal pdFecSis As Date, _
                                        ByVal pnColPage As Integer, ByVal pbValido As Boolean, _
                                        ByVal psNroCH As String, ByVal pnNroProc As Integer, ByVal psCHDesc As String, _
                                        ByVal pdFechaApertReembolso As Date, ByVal pnMontoAsignado As Currency, _
                                        ByVal psPersCod As String, ByVal psPersNombre As String, _
                                        ByVal psCtaArendir As String, ByVal psCtaPendiente As String, _
                                        ByVal psCtaFondoFijo As String, ByVal psSimbolo As String, ByVal pnSaldo As Currency, _
                                        ByVal rsArendir As ADODB.Recordset, ByVal rsEgreDir As ADODB.Recordset) As String
Dim n As Integer
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
Dim Rs As ADODB.Recordset
Dim oCh As NCajaChica
Dim lnTotalSust As Currency


Set Rs = New ADODB.Recordset
Set oCh = New NCajaChica


P = 0
nLin = 66
Cont = 0
BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")

sTexto = PrnSet("MI", 5) & CabeceraRepo(psSimbolo, pnLinPage, psNomCmact, pdFecSis, pnColPage, _
                                         pbValido, psNroCH, pnNroProc, psCHDesc, _
                                         pdFechaApertReembolso, pnMontoAsignado, psPersCod, psPersNombre, nLin, P)
                                         
Linea sTexto, CON & BON & " EGRESOS CON A RENDIR : " & BOFF
Linea sTexto, " --------------------   " & COFF
nLin = nLin + 2
nTotalR = 0:   nTotalI = 0:   nTotalE = 0: nSaldoR = 0: lnTotalSust = 0
If Not rsArendir Is Nothing Then
    Do While Not rsArendir.EOF
        sTexto = sTexto & CabeceraRepo(psSimbolo, pnLinPage, psNomCmact, pdFecSis, pnColPage, _
                                       pbValido, psNroCH, pnNroProc, psCHDesc, _
                                       pdFechaApertReembolso, pnMontoAsignado, psPersCod, psPersNombre, nLin, P)
        nLin = nLin + 2
        Linea sTexto, CON & "  Recibo de A rendir: " & rsArendir![Nro Doc] & "   " & rsArendir!fecha & "   " & _
              Space(15) & "Monto : " & Right(Space(18) & psSimbolo & " " & Format(rsArendir!Importe, "#,#0.00"), 18) & ImpreFormat(IIf(Val(rsArendir!NroCajaCh) = pnNroProc, "", " Caja Chica N° " & rsArendir!NroCajaCh), 20, 10) & COFF
        Linea sTexto, CON & "  Usuario...........: " & rsArendir!Solicitante & COFF
        If Val(rsArendir!NroCajaCh) = pnNroProc Then
            nTotalR = nTotalR + rsArendir!Importe
        End If
        nSaldoR = nSaldoR + rsArendir!Saldo
        'se buscan las sustentaciones
        Set Rs = oCh.GetCHSustAtenArendir(psCtaArendir, psCtaPendiente, rsArendir!cMovNroAtenc)
        If Not Rs.EOF Then
           Do While Not Rs.EOF
              Cont = Cont + 1
              nLin = nLin + 1
              sTexto = sTexto & CabeceraRepo(psSimbolo, pnLinPage, psNomCmact, pdFecSis, pnColPage, _
                                             pbValido, psNroCH, pnNroProc, psCHDesc, _
                                             pdFechaApertReembolso, pnMontoAsignado, psPersCod, psPersNombre, nLin, P)
                                             
              Linea sTexto, CON + PrnSet("I+") & ImpreFormat(Format(Cont, "000"), 3, 5) & ImpreFormat(Rs!FECHADOC, 10) & _
                       ImpreFormat(Rs!cDocAbrev, 3) & ImpreFormat(Rs!cDocNro, 15) & ImpreFormat(PstaNombre(Rs!cPersNombre), 35) & _
                       ImpreFormat(Rs!cMovDesc, 32) & ImpreFormat(Rs!Importe, 12, 2, True) & PrnSet("I-") + COFF
              lnTotalSust = lnTotalSust + Abs(Rs!Importe)
              Rs.MoveNext
           Loop
        End If
        Rs.Close
        Set Rs = Nothing
        If Val(rsArendir!Saldo) = 0 Then
            Set Rs = oCh.GetDatosRendArendirCH(rsArendir!cMovNroAtenc, psCtaFondoFijo)
            If Not Rs.EOF Then
               Cont = Cont + 1
               nLin = nLin + 1
               CabeceraRepo psSimbolo, pnLinPage, psNomCmact, pdFecSis, pnColPage, _
                             pbValido, psNroCH, pnNroProc, psCHDesc, _
                             pdFechaApertReembolso, pnMontoAsignado, psPersCod, psPersNombre, nLin, P
                             
               Linea sTexto, CON + BON + ImpreFormat(Format(Cont, "000"), 3, 5) & ImpreFormat(Rs!FechaRend, 10) & _
                            ImpreFormat(IIf(Rs!NMOVIMPORTE > 0, "R/I ", "R/E"), 3) & ImpreFormat(Rs!cMovNro, 52) & ImpreFormat(IIf(Rs!NMOVIMPORTE > 0, "VUELTO INGRESADO A CAJA", "REEMBOLSO A USUARIO"), 30) & " " & _
                            ImpreFormat(Abs(Rs!NMOVIMPORTE), 13, 2, True) + BOFF + COFF
                            
               If Rs!NMOVIMPORTE > 0 Then
                     nTotalI = nTotalI + Abs(Rs!NMOVIMPORTE)
               Else
                     nTotalE = nTotalE + Abs(Rs!NMOVIMPORTE)
               End If
            End If
            Rs.Close
            Set Rs = Nothing
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
If Not rsEgreDir Is Nothing Then
    Linea sTexto, BON & " EGRESOS DIRECTOS : " & BOFF
    Linea sTexto, " ----------------   " & COFF
    Cont = 0: nTotal = 0
    Do While Not rsEgreDir.EOF
          sTexto = sTexto & CabeceraRepo(psSimbolo, pnLinPage, psNomCmact, pdFecSis, pnColPage, _
                                        pbValido, psNroCH, pnNroProc, psCHDesc, _
                                        pdFechaApertReembolso, pnMontoAsignado, psPersCod, psPersNombre, nLin, P)
          nLin = nLin + 1
          Cont = Cont + 1
          Linea sTexto, CON & "   " & Format(Cont, "000") & " " & Justifica(rsEgreDir!Tipo, 4) & " " & Justifica(rsEgreDir!fecha, 10) & " " & _
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
Private Function CabeceraRepo(ByVal psSimbolo As String, ByVal pnLinPage As Integer, ByVal psNomCmact As String, ByVal pdFecSis As Date, _
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
   Linea sCabe, COFF + ImpreFormat(psNomCmact, 60, 0) & Space(42) & pdFecSis & " - " & Format(Time, "hh:mm:ss")
   Linea sCabe, "     CONTABILIDAD    " & Space(55) & "Pag. " & Format(P, "000")
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
                              ByVal pnColPage As Integer, ByVal pnLinPage As Integer, ByVal pdFecSis As Date, _
                              ByVal psEmpresaLogo As String, ByVal psOpeCod As String, _
                              ByVal psMovNro As String, ByVal pbValido As Boolean, ByVal psCtaFondoFijo As String) As String
Dim sTexto   As String
Dim sAsiento As String
Dim sMovs    As String
Dim sMov     As String
Dim sMov1    As String
Dim sUltMovRend As String
Dim sFecAsiento As String
Dim Rs As ADODB.Recordset
Dim oCh As NCajaChica
Dim oGen As DGeneral
Dim lsImp As String
Dim lnTotalDebe As Currency
Dim lnTotalHaber As Currency
Dim nLi    As Integer, P As Integer
Dim sCabe  As String

Set oGen = New DGeneral
Set Rs = New ADODB.Recordset
Set oCh = New NCajaChica

Linea sAsiento, ImpreCabAsiento(pnColPage, pdFecSis, psEmpresaLogo, psOpeCod, psMovNro, "", True, Format(pdFecSis, gsFormatoFechaView), False, 2)
If pbValido = False Then
   Linea sAsiento, BON & Centra("*** REPORTE NO VALIDO PARA RENDICION EN CONTABILIDAD ***", 80) & BOFF, 2
End If

Set Rs = oCh.GetAsientoCH(psAreaCh, psAgeCh, pnProcCH)
lnTotalDebe = 0
lnTotalHaber = 0
sTexto = ""
Do While Not Rs.EOF
   Linea sTexto, ImpreFormat(Rs!cCtaContCod, 19, 5) & "  " & ImpreFormat(oGen.CuentaNombre(Rs!cCtaContCod), 60) & ImpreFormat(psSimbolo, 3) & ImpreFormat(Rs!Debe, 12, 2, True) & ImpreFormat(psSimbolo, 3) & ImpreFormat(Rs!Haber, 12, 2, True)
   lnTotalDebe = lnTotalDebe + Rs!Debe
   lnTotalHaber = lnTotalHaber + Rs!Haber
   Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing
lsImp = sAsiento & oImpresora.gPrnSaltoLinea & ImpreAsiento(pnColPage, psSimbolo, sTexto, lnTotalDebe, lnTotalHaber, True, "ENCARGADO")

''Ahora Imprimimos Detalle de Documentos por Cuenta en el Debe
Set Rs = oCh.GetDocumentosAsientoCH(psAreaCh, psAgeCh, pnProcCH, psCtaFondoFijo)
sCabe = CON + String(125, "=") & oImpresora.gPrnSaltoLinea _
      & "  D O C U M E N T O               C O N C E P T O                                                DEBE           HABER" & oImpresora.gPrnSaltoLinea _
      & "TIPO  NUMERO        FECHA " & oImpresora.gPrnSaltoLinea _
      & String(125, "-") & oImpresora.gPrnSaltoLinea

nLi = pnLinPage: P = 0
sTexto = PrnSet("MI", 4)
Do While Not Rs.EOF
   If nLi > pnLinPage - 6 Then
      Linea sTexto, COFF + Cabecera("DETALLE DE GASTOS DE CAJA CHICA POR DOCUMENTO", P, psSimbolo, pnColPage, sCabe)
      nLi = 9
   End If
   Linea sTexto, CON + BON + "CUENTA CONTABLE: " & ImpreFormat(Rs!cCtaContCod, 18) & ImpreFormat(oGen.CuentaNombre(Rs!cCtaContCod), 80) + BOFF, 2
   nLi = nLi + 2
   sAsiento = Rs!cCtaContCod
   lnTotalDebe = 0: lnTotalHaber = 0
   Do While sAsiento = Rs!cCtaContCod
      lnTotalDebe = lnTotalDebe + Rs!Debe
      lnTotalHaber = lnTotalHaber + Rs!Haber
      Linea sTexto, ImpreFormat(Rs!cDocAbrev, 3, 0) & ImpreFormat(Rs!cDocNro, 16) & ImpreFormat(Rs!dDocFecha, 10) & ImpreFormat(Rs!cMovDesc, 50) & ImpreFormat(Rs!Debe, 14, 2, True) & ImpreFormat(Rs!Haber, 14, 2, True)
      nLi = nLi + 1
      If nLi > pnLinPage - 6 Then
         Linea sTexto, COFF + Cabecera("DETALLE DE GASTOS DE CAJA CHICA POR DOCUMENTO", P, psSimbolo, pnColPage, sCabe)
         nLi = 9
      End If
      Rs.MoveNext
      If Rs.EOF Then Exit Do
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
                                        ByVal psOpeCod As String, ByVal psMovNro As String, ByVal psNomCmact As String, Optional ByVal lnNumCop As Long = 1)
Dim lsImpre As String
Dim psTit1 As String
Dim Rs As ADODB.Recordset
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

Set Rs = oCaja.GetDatosHabilitacion(psMovNro)
If Not Rs.EOF And Not Rs.BOF Then
    Do While Not Rs.EOF
        If Rs!Origen <> "" Then lsAreaOrigen = Rs!Origen
        If Rs!Destino <> "" Then lsAreaDestino = Rs!Destino
        lsTipoTransporte = Rs!cTipoTrans
        lsNombreTransp = Trim(Rs!cNombreTransp)
        lnImporteTrans = Rs!nMontoTrans
        lsNroDocTransp = Rs!DocCompTransp
        lsNroDocSal = Rs!DocSalEfecBov
        Rs.MoveNext
    Loop
Else
    Exit Function
End If
Rs.Close
Set Rs = Nothing
lsImpre = ""
psTit1 = "COMPROBANTE DE SALIDA EN EFECTIVO  " & IIf(lsNroDocSal = "", "", " N° " & lsNroDocSal)
lsImpre = lsImpre & ImpreCabAsiento(pnColPage, CDate(psFecha), psNomCmact, psOpeCod, psMovNro, psTit1, True) & oImpresora.gPrnSaltoLinea
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
lsBilletaje = ImprimeBilletaje(Mid(psOpeCod, 3, 1), psMovNro, pnColPage, False)
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
                                        ByVal psOpeCod As String, ByVal psMovNro As String, ByVal psNomCmact As String)
Dim lsImpre As String
Dim psTit1 As String
Dim Rs As ADODB.Recordset
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

Set Rs = oCaja.GetDatosConfHabAgencias(psMovNro)
If Not Rs.EOF And Not Rs.BOF Then
    Do While Not Rs.EOF
        If Rs!Origen <> "" Then lsAreaOrigen = Rs!Origen
        If Rs!Destino <> "" Then lsAreaDestino = Rs!Destino
        lsTipoTransporte = Rs!cTipoTrans
        lsNombreTransp = Trim(Rs!cNombreTransp)
        lnImporteTrans = Rs!nMontoTrans
        lsNroDocTransp = Rs!DocCompTransp
        lsNroDocSal = Rs!DocSalEfecBov
        Rs.MoveNext
    Loop
Else
    Exit Function
End If
Rs.Close
Set Rs = Nothing
lsImpre = ""
psTit1 = "CONFIRMACION DE HABILITACION EN EFECTIVO"
lsImpre = lsImpre & ImpreCabAsiento(pnColPage, CDate(psFecha), psNomCmact, psOpeCod, psMovNro, psTit1, True) & oImpresora.gPrnSaltoLinea
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
lsBilletaje = ImprimeBilletaje(Mid(psOpeCod, 3, 1), psMovNro, pnColPage, False)
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
Print #nFicSal, oImpresora.gPrnTamPaginaCab + Chr$(lnLineasBol);  'Longitud de página a 22 líneas'
Print #nFicSal, oImpresora.gPrnTamLetra10CPI;   'Tamaño 10 cpi
Print #nFicSal, oImpresora.gPrnTpoLetraSansSerif;     'Tipo de Letra Sans Serif
Print #nFicSal, oImpresora.gPrnCondensadaOFF  ' cancela condensada
Print #nFicSal, oImpresora.gPrnBoldOFF  ' desactiva negrita

lsFecha = Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4)
'20010703 1347191120100EJRS
lsHora = Format$(Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2), "hh:mm:ss")
lsUserOrig = Right(psMovNro, 4)
lnEspaceIni = 2
lnLineasBol = 22
lnAnchoBol = 50
Print #nFicSal, lsNegritaOn; 'Activa Negrita
Print #nFicSal, ImpreFormat("CMACT - AHORROS", lnAnchoBol, lnEspaceIni) & ImpreFormat("CMACT - AHORROS", lnAnchoBol, lnEspaceIni)
lnNumLineas = lnNumLineas + 1
'''Print #nFicSal, ImpreFormat(Trim(psNomAge) + IIf(pnMoneda = gMonedaNacional, "-SOLES", "-DOLARES"), lnAnchoBol, lnEspaceIni) & ImpreFormat(Trim(psNomAge) + IIf(pnMoneda = gMonedaNacional, "-SOLES", "-DOLARES"), lnAnchoBol, lnEspaceIni) 'MARG ERS044-2016
Print #nFicSal, ImpreFormat(Trim(psNomAge) + IIf(pnMoneda = gMonedaNacional, "-" & StrConv(gcPEN_PLURAL, vbUpperCase), "-DOLARES"), lnAnchoBol, lnEspaceIni) & ImpreFormat(Trim(psNomAge) + IIf(pnMoneda = gMonedaNacional, "-" & StrConv(gcPEN_PLURAL, vbUpperCase), "-DOLARES"), lnAnchoBol, lnEspaceIni) 'MARG ERS044-2016
lnNumLineas = lnNumLineas + 1
Print #nFicSal, ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni)
lnNumLineas = lnNumLineas + 1
Print #nFicSal, ""
lnNumLineas = lnNumLineas + 1
Print #nFicSal, ImpreFormat("PARA :[" & psUserDest & "] " & psNomUserDest, lnAnchoBol, lnEspaceIni) & ImpreFormat("PARA :[" & psUserDest & "] " & psNomUserDest, lnAnchoBol, lnEspaceIni)
lnNumLineas = lnNumLineas + 1
Print #nFicSal, ImpreFormat("DE   :[" & psUserOrig & "] " & psNomUserOrigen, lnAnchoBol, lnEspaceIni) & ImpreFormat("DE   :[" & psUserOrig & "] " & psNomUserOrigen, lnAnchoBol, lnEspaceIni)
lnNumLineas = lnNumLineas + 1
Print #nFicSal, ""
lnNumLineas = lnNumLineas + 1
Print #nFicSal, CentrarCadena(psTitCab, lnAnchoBol - 10) & CentrarCadena(psTitCab, lnAnchoBol)
lnNumLineas = lnNumLineas + 1
Print #nFicSal, ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni) & ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni)
lnNumLineas = lnNumLineas + 1
Print #nFicSal, ImpreFormat(ImpreFormat(psTitDet, lnAnchoBol / 2) & ImpreFormat(CCur(pnImporte), 15, 2, True), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat(psTitDet, lnAnchoBol / 2) & ImpreFormat(CCur(pnImporte), 15, 2, True), lnAnchoBol, lnEspaceIni)
lnNumLineas = lnNumLineas + 1
For I = lnNumLineas To lnNumLineas + 3
    Print #nFicSal, ""
Next
lnNumLineas = I
Print #nFicSal, ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni) & ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni)
lnNumLineas = lnNumLineas + 1
Print #nFicSal, ImpreFormat(lsUserOrig, lnAnchoBol, lnEspaceIni) & ImpreFormat(lsUserOrig, lnAnchoBol, lnEspaceIni)
lnNumLineas = lnNumLineas + 1

Dim oGen As DGeneral
Set oGen = New DGeneral
lsMensaje = oGen.GetMensajeBoletas("")
Set oGen = Nothing
Print #nFicSal, ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) & ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni)
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
Dim sOpe As String, nOPE As Integer
Dim sCta As String, nCta As Integer
Dim sObj As String, nObj As Integer
Dim sDoc As String, nDoc As Integer
Dim sSql As String, nLin As Integer, nPag As Integer
Dim I As Integer
Dim clsOpe As New DOperacion
Dim clsCta As New DCtaCont
Dim sTexto As String
Dim mMat(4000, 12) As String

On Error GoTo ImprimeOperacionesErr


Linea sTexto, ""
Set Rope = clsOpe.CargaOpeTpo("")
If Not Rope.EOF Then
   RaiseEvent BarraShow(Rope.RecordCount)
   nOPE = 0
   Do While Not Rope.EOF
      RaiseEvent BarraProgress(Rope.Bookmark, "OPERACIONES", "", "Procesando... ", vbBlue)
      nOPE = nOPE + 1
      mMat(nOPE, 0) = Format(Rope!nOpeNiv, "000")
      mMat(nOPE, 1) = Rope!COPECOD
      mMat(nOPE, 2) = Rope!cOpeDesc
      sOpe = Rope!COPECOD
      
      Set Rdoc = clsOpe.CargaOpeDoc(sOpe)
      
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
                  mMat(nObj, 7) = Robj!cObjetoCod
                  mMat(nObj, 8) = Robj!cObjetoDesc
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
   RaiseEvent BarraClose
End If

nLin = 1
nPag = 1
sTexto = ""
For I = 1 To nOPE
   'DoEvents
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
Set clsCta = Nothing
Set clsOpe = Nothing
RSClose Robj
RSClose Rope
RSClose Rcta
ImprimeOperaciones = PrnSet("C+") & sTexto & PrnSet("C-")
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

Public Function ImprimeRepCtaCol(pnLinPage As Integer) As String
Dim Rs As ADODB.Recordset
Dim lsOperacion As String
Dim lnColumna As Integer

Dim sImpre As String
Dim nLin As Integer
Dim nPageNro As Integer
Dim clsRepCtaCol As DRepCtaColumna
Set clsRepCtaCol = New DRepCtaColumna
nLin = pnLinPage
sImpre = ""
On Error GoTo ImprimeRepCtaColErr
Set Rs = clsRepCtaCol.CargaRepImpresion
RaiseEvent BarraShow(Rs.RecordCount)
   Do While Not Rs.EOF
      RaiseEvent BarraProgress(Rs.Bookmark, "Reporte con Cuentas por Columnas", "", "Proceando... ", vbBlue)
      If nLin > pnLinPage - 4 Then
         Linea sImpre, CabeRepo(vsEmpNom, vsAgeNom, "Contabilidad", "", vsFecSis, "Configuración de Reportes de Cuentas Contables por Columnas", _
               String(65, "-"), "", "", nPageNro, 79), 0
         nLin = 6
      End If
      If lsOperacion <> Trim(Rs!COPECOD) Then
          Linea sImpre, ""
          Linea sImpre, BON & Trim(Rs!COPECOD) & " - " & Trim(Rs!cOpeDesc) & " " & IIf(Mid(Rs!COPECOD, 3, 1) = "1", "MN", "ME") & " (Anexo : " & Trim(Right(Rs!COPECOD, 1)) & ")" & BOFF
          nLin = nLin + 2
      End If
      If lnColumna <> Rs!nNroCol Then
          Linea sImpre, Space(5) & Rs!nNroCol & "  " & Trim(Rs!cDescCol)
          nLin = nLin + 1
      End If
      Linea sImpre, Space(15) & Rs!cCtaContCod & " " & Trim(Rs!cCtaContDesc)
      nLin = nLin + 1
      lnColumna = Trim(Rs!nNroCol)
      lsOperacion = Trim(Rs!COPECOD)
      Rs.MoveNext
    Loop
RaiseEvent BarraClose
RSClose Rs
Set clsRepCtaCol = Nothing
ImprimeRepCtaCol = sImpre
Exit Function
ImprimeRepCtaColErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeRepCtaCol Method")
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
RaiseEvent BarraShow(prs.RecordCount)
Do While Not prs.EOF
   RaiseEvent BarraProgress(prs.Bookmark, "PLAN DE CUENTAS", "", "Procesando... ", vbBlue)
  'DoEvents
   ' Sangra Código
   nLongT = Len(Trim(prs!cCtaContCod))
   If nLongC < nLongT Then
      Espacios = Espacios + nLongT
      aSpace(nLongT, 1) = nLongT
      aSpace(nLongT, 2) = Espacios
      If nLongT <= 6 Then
         If lnPagNro > 0 Then
            Linea sTexto, ""
            lnLinea = lnLinea + 1
         End If
      End If
   End If
   If nLongC > nLongT Then
      Espacios = aSpace(nLongT, 2)
      nNivel = nLongT
      If nNivel <= 6 And lnPagNro > 0 Then
         Linea sTexto, ""
         lnLinea = lnLinea + 1
      End If
   End If
   nLongC = Len(Trim(prs!cCtaContCod))
   nLongD = Len(Trim(prs!cCtaContDesc))
   
   'Salto de Pagina si cumple condicion
   If lnLinea > pnLinPage - lnPiePag Then
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
RaiseEvent BarraClose
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
RaiseEvent BarraShow(prs.RecordCount)
Do While Not prs.EOF
   RaiseEvent BarraProgress(prs.Bookmark, "OBJETOS", "", "Procesando... ", vbBlue)
   If lnLinea > lnCarLin - 2 Then
      Linea lsImpre, Cabecera("REPORTE DE OBJETOS", lnPageNro, "", pnLinPage), 0
      Linea lsImpre, " CODIGO        DESCRIPCION"
      Linea lsImpre, String(79, "-")
      lnLinea = 6
   End If
   Linea lsImpre, Space(2) & Trim(prs!cObjetoCod) _
               & Space(5) & Trim(prs!cObjetoDesc)
   lnLinea = lnLinea + 1
   prs.MoveNext
Loop
RaiseEvent BarraClose
prs.Close: Set prs = Nothing
Linea lsImpre, String(lnCarLin, "-")
ImprimeObjetos = lsImpre
End Function


'***********************************
'***** NUEVAS FUNCIONES
'***********************************
Public Function ImprimeLibroDiario(ByVal dFecIni As Date, ByVal dFecFin As Date, psAge As String, psOpe As String, pnLinPage As Integer) As String
Dim sSql As String
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

    On Error GoTo ImprimeLibroDiarioErr
    nTotalDocD = 0
    nTotalDocH = 0
    nTotalD = 0: nTotalH = 0
    sCadRep = ""
    P = 0
    Dim oFun As New NContFunciones
    Set R = oFun.GetLibroDiario(dFecIni, dFecFin, psAge, psOpe)
    Set oFun = Nothing
    If R.EOF Then
       Err.Raise 50001, "NContImprimir", "No se registraron Asientos para generar Libro Diario"
       R.Close: Set R = Nothing
       Exit Function
    End If
      sCadRep = PrnSet("MI", 8)
      nLi = pnLinPage
      RaiseEvent BarraShow(R.RecordCount)
      sClave = ""
      sGlosa = R!cOpeGruDesc

      Do While Not R.EOF
        'DoEvents
         RaiseEvent BarraProgress(R.Bookmark, "LIBRO DIARIO", "", "Procesando... ", vbBlue)
         sFec = Mid(R!cMovNro, 7, 2) & "/" & Mid(R!cMovNro, 5, 2) & "/" & Mid(R!cMovNro, 1, 4)
         sAge = Mid(R!cMovNro, 18, 2)         'Agencia
         sTpo = R!cOpeGruCod                  'Tpo de Asiento
         If sClave <> sFec & sAge & sTpo Then
            Linea sCadRep, ImprimeLibroDiarioCab(nLi, P, dFecIni, dFecFin, R!nMovCorrela, sFec, sAge, sTpo, pnLinPage), 0
            If sClave <> "" Then
               Linea sCadRep, BON & Space(5) + "Total de Documentos: " + Space(60) & PrnVal(nTotalDocD, 16, 2) + "  " & PrnVal(nTotalDocH, 16, 2) & BOFF
               Linea sCadRep, BON & Space(5) + "Glosa : " + sGlosa & BOFF, 3
               Linea sCadRep, BON & "Documento: " & Format(R!nMovCorrela, "#####0") & "     Tipo de Asiento: " & R!cOpeGruCod & "      Agencia: " & sAge & "      Fecha: " & sFec & BOFF, 2
               sGlosa = R!cOpeGruDesc
               nLi = nLi + 7
               nTotalDocH = 0
               nTotalDocD = 0
            End If
            sClave = sFec & sAge & sTpo
         End If
         Linea sCadRep, ImprimeLibroDiarioCab(nLi, P, dFecIni, dFecFin, R!nMovCorrela, sFec, sAge, sTpo, pnLinPage), 0
            
         ' Linea de Detalle
         If R!NMOVIMPORTE > 0 Then
            Linea sCadRep, Space(5) + Justifica(R!cCtaContCod, 20) & " " & Justifica(R!cCtaContDesc, 60) & PrnVal(R!NMOVIMPORTE, 16, 2)
            nTotalDocD = nTotalDocD + R!NMOVIMPORTE
            nTotalD = nTotalD + R!NMOVIMPORTE
         Else
            Linea sCadRep, Space(5) + Justifica(R!cCtaContCod, 20) & " " & Justifica(R!cCtaContDesc, 60) & Space(18) & PrnVal(R!NMOVIMPORTE * -1, 16, 2)
            nTotalDocH = nTotalDocH + (R!NMOVIMPORTE * -1)
            nTotalH = nTotalH + (R!NMOVIMPORTE * -1)
         End If
         nLi = nLi + 1
         If P Mod 20 = 0 Then
            sCad = sCad & sCadRep
            sCadRep = ""
         End If
         R.MoveNext
      Loop
      R.MoveLast
      Linea sCadRep, ImprimeLibroDiarioCab(nLi, P, dFecIni, dFecFin, R!nMovCorrela, sFec, sAge, sTpo, pnLinPage), 0
      Linea sCadRep, BON & Space(5) + "Total de Documentos: " + Space(60) & PrnVal(nTotalDocD, 16, 2) + "  " & PrnVal(nTotalDocH, 16, 2) & BOFF
      Linea sCadRep, BON & Space(5) + "Glosa : " + sGlosa & BOFF, 3
      Linea sCadRep, BON & "     T O T A L E S  " & Space(66) & PrnVal(nTotalD, 16, 2) & "  " & PrnVal(nTotalH, 16, 2) & BOFF
      sCad = sCad & sCadRep
      R.Close: Set R = Nothing
    ImprimeLibroDiario = sCad
    RaiseEvent BarraClose
    
    Exit Function
ImprimeLibroDiarioErr:
    RaiseEvent BarraClose
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeLibroDiario Method")
End Function

Private Function ImprimeLibroDiarioCab(nLi As Integer, P As Integer, dFecIni As Date, dFecFin As Date, nNro As Currency, sFec As String, sAge As String, sTpo As String, pnLinPage As Integer) As String
Dim sCadCab As String
If pnLinPage - 6 < nLi Then
   Linea sCadCab, COFF & BON & Cabecera("L I B R O   D I A R I O ", P, "", 80, "", Format(dFecFin, gsFormatoFechaView)), 2
   Linea sCadCab, CON & Space(5) + "CODIGO" + Space(15) + "NOMBRE" + Space(64) + "DEBE" + Space(13) + "HABER", 2
   Linea sCadCab, "Documento: " & Format(nNro, "#####0") & "     Tipo de Asiento: " & sTpo & "      Agencia: " & sAge & "      Fecha: " & sFec & BOFF, 2
   nLi = 9
End If
ImprimeLibroDiarioCab = sCadCab
End Function

Public Function ImprimeLibroMayor(dFechaDel As Date, dFechaAl As Date, psAge As String, psOpe As String, pnLinPage As Integer) As String
Dim rsCta As New ADODB.Recordset
Dim nImporte As Currency
Dim nSaldo As Currency, nSaldoIni As Currency
Dim nDebe  As Currency, nHaber  As Currency
Dim nDebeD As Currency, nHaberH As Currency
Dim nDebeT As Currency, nHaberT As Currency
Dim nHaberD As Currency
Dim sTexto As String, sImpre As String
Dim sFecha As String
Dim sSql   As String
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
sSql = "SELECT a.cCtaContCod, c.cCtaContDesc, m1.nMovCorrela, b.cMovNro, OT.cOpeGruCod, cc.cCtaCaracter, " _
     & "       ISNULL(SUM(CASE WHEN a.nMovImporte > 0 THEN a.nMovImporte END),0) as nDebe, " _
     & "       ISNULL(SUM(CASE WHEN a.nMovImporte < 0 THEN a.nMovImporte * -1 END),0) as nHaber " _
     & "FROM   Mov b    JOIN MovCta a ON a.nMovNro = b.nMovNro JOIN " & vsBaseComunes & "CtaCont c ON c.cCtaContCod = a.cCtaContCod" _
     & "                JOIN MovCont m1 ON m1.nMovNro = b.nMovNro " _
     & "                JOIN " & vsBaseComunes & "OpeTpo OT ON OT.cOpeCod = b.cOpeCod " _
     & "                JOIN " & vsBaseComunes & "CtaContClase cc ON c.cCtaContCod LIKE cc.cCtaContCod + '%' " _
     & "WHERE  b.nMovEstado = " & gMovEstContabMovContable & " and b.nMovFlag <> " & gMovFlagEliminado & " and " _
     & "       substring(b.cmovnro,1,8) BETWEEN '" & Format(dFechaDel, "yyyymmdd") & "' and '" & Format(dFechaAl, "yyyymmdd") & "' " _
     & "GROUP BY a.cCtaContCod, c.cCtaContDesc, m1.nMovCorrela, b.cMovNro, OT.cOpeGruCod, cc.cCtaCaracter ORDER BY a.cCtaContCod, m1.nMovCorrela, SubString(b.cMovNro,1,8), Substring(b.cMovNro,18,2), OT.cOpeGruCod"
If oConec.AbreConexion = False Then Exit Function
Set rsCta = oConec.CargaRecordSet(sSql)
If rsCta.EOF Then
   oConec.CierraConexion
   Set oConec = Nothing
   Err.Raise 50001, "No se registraron movimientos ...!"
   Exit Function
End If

RaiseEvent BarraShow(rsCta.RecordCount)
sTexto = PrnSet("MI", 5)
sImpre = ""
nDebeT = 0: nHaberT = 0
Set oSdo = New NCtasaldo
Do While Not rsCta.EOF
   'Obtenemos el Saldo de la Cuenta al día dFechaDel
   sCtaCod = rsCta!cCtaContCod
   sCtaDes = rsCta!cCtaContDesc
   nSaldo = oSdo.GetCtaSaldo(sCtaCod, Format(dFechaDel - 1, gsFormatoFecha))
   Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage), 0
   nDebe = 0: nHaber = 0
   nDebeD = 0: nHaberD = 0
   If rsCta.EOF Then
      Exit Do
   End If
   sFecha = Mid(rsCta!cMovNro, 1, 8)

   sTipoCta = RTrim(rsCta!cCtaCaracter)

   Do While rsCta!cCtaContCod = sCtaCod
     'DoEvents
      RaiseEvent BarraProgress(rsCta.Bookmark, "LIBRO MAYOR", "", "Procesando... ", vbBlue)
      If P Mod 20 = 0 Then
         sImpre = sImpre & sTexto
         sTexto = ""
      End If
      If sFecha <> Mid(rsCta!cMovNro, 1, 8) Then
         Linea sTexto, BON & Space(15) & "Total Diario  : " & PrnVal(nDebeD, 16, 2, False) & "  " & PrnVal(nHaberD, 16, 2, False) & "  " & PrnVal(nSaldo, 16, 2) & BOFF, 2
         nLin = nLin + 2
         nDebeD = 0: nHaberD = 0
         sFecha = Mid(rsCta!cMovNro, 1, 8)
      End If
      Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage), 0
      If sTipoCta = "D" Then
         nSaldo = nSaldo + rsCta!nDebe - rsCta!nHaber
      Else
         nSaldo = nSaldo + rsCta!nHaber - rsCta!nDebe
      End If
      nSaldoIni = nSaldo
      nDebe = nDebe + rsCta!nDebe
      nHaber = nHaber + rsCta!nHaber
      nDebeD = nDebeD + rsCta!nDebe
      nHaberD = nHaberD + rsCta!nHaber
      nDebeT = nDebeT + rsCta!nDebe
      nHaberT = nHaberT + rsCta!nHaber

      'Linea Detalle de Reporte
      Linea sTexto, Mid(rsCta!cMovNro, 7, 2) & "/" & Mid(rsCta!cMovNro, 5, 2) & "/" & Mid(rsCta!cMovNro, 1, 4) & " " & PrnVal(rsCta!nMovCorrela, 5, 0) & "   " & Mid(rsCta!cMovNro, 18, 2) & "-" _
             & rsCta!cOpeGruCod & Space(7) & PrnVal(rsCta!nDebe, 16, 2, False) & "  " _
             & PrnVal(rsCta!nHaber, 16, 2, False) & "  " _
             & PrnVal(nSaldo, 16, 2)
      nLin = nLin + 1
      rsCta.MoveNext
      If rsCta.EOF Then
         Exit Do
      End If
   Loop
   Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage), 0
   Linea sTexto, BON & Space(15) & "Total Diario  : " & PrnVal(nDebeD, 16, 2, False) & "  " & PrnVal(nHaberD, 16, 2, False) & "  " & PrnVal(nSaldo, 16, 2) & BOFF, 2
   nLin = nLin + 2
   Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage), 0
   Linea sTexto, BON & Space(15) & "Total Cuenta  : " & PrnVal(nDebe, 16, 2, False) & "  " & PrnVal(nHaber, 16, 2, False) & "  " & PrnVal(nSaldo, 16, 2) & BOFF, 2
   nLin = nLin + 2
   Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage) & BON, 0
   Linea sTexto, Space(15) & "SALDO INICIAL : " & PrnVal(nSaldoIni, 16, 2)
   Linea sTexto, Space(15) & "SALDO FINAL   : " & PrnVal(nSaldo, 16, 2) & BOFF
   Linea sTexto, "---------------------------------------------------------------------------------------", 2
   nLin = nLin + 4

   If Not rsCta.EOF Then
      If nLin > pnLinPage - 8 Then
         sCtaCod = rsCta!cCtaContCod
         sCtaDes = rsCta!cCtaContDesc
         Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage), 0
      Else
         Linea sTexto, BON & "CUENTA: " & Justifica(rsCta!cCtaContCod, 20) & " " & rsCta!cCtaContDesc & BOFF, 2
         nLin = nLin + 2
      End If
   End If
Loop
Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage), 0
Linea sTexto, BON & Space(15) & "TOTALES       : " & PrnVal(nDebeT, 16, 2, False) & "  " & PrnVal(nHaberT, 16, 2, False) & BOFF, 2
nLin = nLin + 2
ImprimeLibroMayor = sTexto

   RaiseEvent BarraClose
   oConec.CierraConexion
   Set oConec = Nothing
   Set oSdo = Nothing

    Exit Function
ImprimeLibroMayorErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeLibroMayor Method")
End Function

Private Function ImprimeLibroMayorCab(ByRef nLin As Integer, ByRef P As Integer, dFechaDel As Date, dFechaAl As Date, sCtaCod As String, sCtaDes As String, pnLinPage As Integer) As String
Dim k As Integer, nObjs As Integer
Dim sCabe As String
If nLin > pnLinPage - 6 Then
   Linea sCabe, BON & Cabecera(" L I B R O    M A Y O R ", P, "", 80, , Format(dFechaAl, gsFormatoFechaView)), 0
   Linea sCabe, BON & Centra("Del " & dFechaDel & " al " & dFechaAl, 80), 2
   Linea sCabe, "FECHA       DOC   ASIENTO                  DEBE           HABER      SALDO FINAL", 3
   Linea sCabe, "CUENTA: " & Justifica(sCtaCod, 20) & " " & sCtaDes & BOFF, 2
   nLin = 10
End If
ImprimeLibroMayorCab = sCabe
End Function


Public Function ImprimeMayorCta(psCtaCod As String, psCtaDes As String, pdFechaDel As Date, pdFechaAl As Date, pnImporte As Currency, psFiltro As String, pnLinPage As Integer) As String
Dim sTexto As String
Dim nRow As Integer
Dim sMov As String
Dim sDoc As String
Dim n As Integer
Dim nDebe  As Currency, nHaber  As Currency
Dim nDebeD As Currency, nHaberH As Currency
Dim nHaberD As Currency
Dim nSaldo  As Currency, nSaldoIni As Currency
Dim sFecha  As String
Dim Rs      As ADODB.Recordset
Dim rsCta   As ADODB.Recordset
Dim nLin    As Integer
Dim P       As Integer
Dim sTipoCta As String

nLin = pnLinPage
P = 0
Dim oCont As New NContAsientos
Dim oSdo  As New NCtasaldo
nSaldo = oSdo.GetCtaSaldo(psCtaCod, Format(pdFechaDel - 1, gsFormatoFecha))
nSaldoIni = nSaldo
Set oSdo = Nothing

Set rsCta = oCont.GetMayorCuenta(psCtaCod, Format(pdFechaDel, "yyyymmdd"), Format(pdFechaAl, "yyyymmdd"), pnImporte, psFiltro)
Set oCont = Nothing
If rsCta.EOF Then
   Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
   Linea sTexto, BON & "CUENTA CONTABLE : " & Justifica(psCtaCod, 16) & " " & psCtaDes, 2
   Linea sTexto, CON & Space(87) & "SALDO INICIAL : " & PrnVal(nSaldo, 16, 2)
   Linea sTexto, Space(87) & "SALDO FINAL   : " & PrnVal(nSaldo, 16, 2) & COFF & BOFF, 1
   ImprimeMayorCta = sTexto
   Set oCont = Nothing
   Exit Function
End If
Dim oCta  As New DCtaCont

Set Rs = oCta.CargaCtaContClase(psCtaCod)
Set oCta = Nothing
If Not Rs.EOF Then
   sTipoCta = Trim(Rs!cCtaCaracter)
Else
   sTipoCta = "D"
End If

RaiseEvent BarraShow(rsCta.RecordCount)
sTexto = ""
Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
Linea sTexto, BON & "CUENTA CONTABLE : " & Justifica(psCtaCod, 16) & " " & psCtaDes & BOFF, 2
nLin = nLin + 2
nDebe = 0: nHaber = 0
n = 1

sFecha = Mid(rsCta!cMovNro, 1, 8)
Do While Not rsCta.EOF
  'DoEvents
   RaiseEvent BarraProgress(rsCta.Bookmark, "Mayor de Cuenta Contable : " & psCtaCod, "", "Generando Reporte... ", vbBlue)
   sMov = rsCta!cMovNro
   If sFecha <> Mid(sMov, 1, 8) Then
      Linea sTexto, CON & BON & Space(87) & "TOTAL DIARIO  : " & PrnVal(nDebeD, 16, 2, False) & "  " & PrnVal(nHaberD, 16, 2, False) & "  " & PrnVal(nSaldo, 16, 2) & COFF & BOFF, 2
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
   sDoc = Justifica(rsCta!cDocAbrev, 3) & " " & Justifica(rsCta!cDocNro, 20) & " " & rsCta!dDocFecha & " "
   sMov = rsCta!cMovNro
   Linea sTexto, CON & " " & Mid(sMov, 1, 8) & "-" & Mid(sMov, 9, 6) & " " & Right(RTrim(sMov), 4) & "  " _
          & " " & sDoc & Mid(Replace(Replace(rsCta!cMovDesc, Chr(13), " "), oImpresora.gPrnSaltoLinea, "") & Space(100), 1, 42) & " " _
          & PrnVal(rsCta!nDebe, 16, 2, False) & "  " & PrnVal(rsCta!nHaber, 16, 2, False) _
          & COFF
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
      Linea sTexto, CON & " " & Space(23) & sDoc & "  " & COFF
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
Linea sTexto, CON & BON & Space(87) & "TOTAL DIARIO  : " & PrnVal(nDebeD, 16, 2, False) & "  " & PrnVal(nHaberD, 16, 2, False) & "  " & PrnVal(nSaldo, 16, 2) & COFF & BOFF, 2
nLin = nLin + 2
n = n + 1
Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
Linea sTexto, CON & " -------------------------" & String(36, "-") & "------------------------------------------------------------------------------------------------"
Linea sTexto, BON & Space(87) & "TOTAL CUENTA  : " & PrnVal(nDebe, 16, 2, False) & "  " & PrnVal(nHaber, 16, 2, False) & BOFF, 2
nLin = nLin + 2

Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
Linea sTexto, Space(87) & "SALDO INICIAL : " & PrnVal(nSaldoIni, 16, 2)
Linea sTexto, Space(87) & "SALDO FINAL   : " & PrnVal(nSaldo, 16, 2)
Linea sTexto, COFF & BOFF
nLin = pnLinPage + 1
Linea sTexto, CON & " =========================" & String(36, "=") & "================================================================================================" & COFF
ImprimeMayorCta = sTexto
RaiseEvent BarraClose
Rs.Close: Set Rs = Nothing
rsCta.Close: Set rsCta = Nothing

End Function

Private Function ImprimeMayorCuentaCab(ByRef nLin As Integer, ByRef P As Integer, dFechaDel As Date, dFechaAl As Date, sCtaCod As String, sCtaDes As String, pnLinPage As Integer) As String
Dim k As Integer, nObjs As Integer
Dim sCabe As String
Dim oConec As New DConecta

If nLin > pnLinPage - 6 Then
   oConec.AbreConexion
   Linea sCabe, BON & Cabecera(" M A Y O R   D E   C U E N T A S    C O N T A B L E S", P, "", 80, , Format(oConec.GetFechaHoraServer(), gsFormatoFechaView)), 0
   oConec.CierraConexion
   Linea sCabe, BON & Centra("Del " & dFechaDel & " al " & dFechaAl, 80), 2
   Linea sCabe, " ======================" & String(36, "=") & "===================================================================================================="
   Linea sCabe, "  Número               " & "     DOCUMENTO                       " & "                                                          MOVIMIENTO                         "
   Linea sCabe, "  de                   " & " ------------------------------------" & " DESCRIPCION                                     --------------------------        S A L D O "
   Linea sCabe, "  Movimiento           " & " Tipo  Número              Fecha     " & "                                                       DEBE          HABER                   "
   Linea sCabe, " ----------------------" & String(36, "-") & "----------------------------------------------------------------------------------------------------" & COFF
   nLin = 10
End If
ImprimeMayorCuentaCab = sCabe
Set oConec = Nothing
End Function

Public Function ImprimeAsientoConsolidado(psMovNro As String, psMoneda As String, pdFecha As Date, Optional psCtaNoConsideradas As String = "", Optional psComentario As String = "") As String
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lnItem As Long
    Dim lsCadena As String
    Dim lsCabecera As String
    
    Dim CON As String
    Dim COFF As String
    Dim BON As String
    Dim BOFF As String
    
    Dim nTot As Currency
    Dim nTotH As Currency
    
    Dim oDGeneral As DGeneral
    Set oDGeneral = New DGeneral
    
    CON = PrnSet("C+")
    COFF = PrnSet("C-")
    BON = PrnSet("B+")
    BOFF = PrnSet("B-")
    
    Sql = " Select A.cCtaContCod , B.cCtaContdesc + ' ' + C.cCtaContDesc, A.nMovImporte From" _
        & " (" _
        & "     Select cCtaContCod, Sum(nMovImporte) nMovImporte From MovCta MC" _
        & "         Where nMovNro in ('" & psMovNro & "') Group By cCtaContCod" _
        & " ) As A" _
        & " Join Ctacont C On A.cCtaContCod = C.cCtaContCod" _
        & " Join CtaCont B On B.cCtaContCod = LEFT(A.cCtaContCod,LEN(A.cCtaContCod)-2)" _
        & " Where LEFT(A.cCtaContCod,2) Not In ('" & psCtaNoConsideradas & "')" _
        & " order by A.cCtaContCod"
        
    oCon.AbreConexion
    
    lsCadena = " A S I E N T O  C O N T A B L E"
    lsCadena = lsCadena & ImpreCabAsiento(80, pdFecha, "", "__" & psMoneda, "", lsCadena, True) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & JustificaTexto(psComentario, 80, 5) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea

    lsCabecera = ""
    lsCabecera = lsCabecera & CON
    lsCabecera = lsCabecera & "====================================================================================================================================" & oImpresora.gPrnSaltoLinea
    lsCabecera = lsCabecera & "Itm C U E N T A   C O N T A B L E                                                                           DEBE            HABER" & oImpresora.gPrnSaltoLinea
    lsCabecera = lsCabecera & "    Codigo              Descripcion" & oImpresora.gPrnSaltoLinea
    lsCabecera = lsCabecera & "------------------------------------------------------------------------------------------------------------------------------------" & COFF & oImpresora.gPrnSaltoLinea
    
    lsCadena = lsCadena & lsCabecera
    
    Set Rs = oCon.CargaRecordSet(Sql)
    lnItem = 0
    
    While Not Rs.EOF
       lnItem = lnItem + 1
       lsCadena = lsCadena & CON & Format(lnItem, "000") & " " & Justifica(Rs!cCtaContCod, 20) & " " _
                & Justifica(oDGeneral.CuentaNombre(Rs!cCtaContCod), 71) & " " _
                & IIf(Rs!NMOVIMPORTE > 0, PrnVal(Rs!NMOVIMPORTE, 16, 2), Space(16)) & " " _
                & IIf(Rs!NMOVIMPORTE < 0, PrnVal(Rs!NMOVIMPORTE * -1, 16, 2), Space(16)) _
                & COFF & oImpresora.gPrnSaltoLinea
                
        If Rs!NMOVIMPORTE > 0 Then
           nTot = nTot + Val(Rs!NMOVIMPORTE)
        Else
           nTotH = nTotH + Val(Rs!NMOVIMPORTE) * -1
        End If
                
        If lnItem Mod 50 = 0 Then
            lsCadena = lsCadena & CON & "====================================================================================================================================" & COFF & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
            lsCadena = lsCadena & lsCabecera
        End If
        
        Rs.MoveNext
    Wend
    
    If nTot > 0 Or nTotH > 0 Then
       lsCadena = lsCadena & CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
       lsCadena = lsCadena & BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF + oImpresora.gPrnSaltoLinea
    End If
    If Rs.State = adStateOpen Then Rs.Close: Set Rs = Nothing
    lsCadena = lsCadena & CON & "====================================================================================================================================" & COFF & oImpresora.gPrnSaltoLinea
    
    lsCadena = lsCadena & ImprePiePag(80)
    
    lsCadena = lsCadena & "   " & oImpresora.gPrnSaltoLinea

    ImprimeAsientoConsolidado = lsCadena
End Function

Public Function ImprimeAsientoConsolidadoPalcial(psMovNro As String, psMoneda As String, pdFecha As Date, Optional psCtaConsideradas As String = "", Optional psComentario As String = "") As String
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset
    Dim rsD As ADODB.Recordset
    Set rsD = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lnItem As Long
    Dim lsCadena As String
    Dim lsCabecera As String
    
    Dim CON As String
    Dim COFF As String
    Dim BON As String
    Dim BOFF As String
    
    Dim nTot As Currency
    Dim nTotH As Currency
    
    Dim lsNros As String
    
    Dim oDGeneral As DGeneral
    Set oDGeneral = New DGeneral
    
    Dim sqlAdd As String
    Dim lsCtaNoAfecta As String
    
    CON = PrnSet("C+")
    COFF = PrnSet("C-")
    BON = PrnSet("B+")
    BOFF = PrnSet("B-")
    
    Sql = " Select Distinct cCtaContCod From MovCta MC" _
        & " Where nMovNro in ('" & psMovNro & "') And LEft(cCtaContCod,2) IN ('19') And cCtaContCod <> '191403' order by cCtaContCod"
    oCon.AbreConexion
    
    Set Rs = oCon.CargaRecordSet(Sql)
    
    
    
    While Not Rs.EOF
        If Rs!cCtaContCod = "19110709" Then
            sqlAdd = Rs!cCtaContCod & "','191403"
        Else
            sqlAdd = Rs!cCtaContCod
        End If
        
        If Rs!cCtaContCod = "19110708" Then
            lsCtaNoAfecta = "181302"
        ElseIf Rs!cCtaContCod = "19110709" Then
            lsCtaNoAfecta = "181301"
        End If
        
        Sql = " Select nMovNro From MovCta " _
            & " Where nMovNro in ('" & psMovNro & "') And cCtaContCod in ('" & sqlAdd & "')"
        Set rsD = oCon.CargaRecordSet(Sql)
        
        lsNros = ""
        
        While Not rsD.EOF
            If lsNros = "" Then
                lsNros = Trim(Str(rsD!nMovNro))
            Else
                lsNros = lsNros & "','" & Trim(Str(rsD!nMovNro))
            End If
            rsD.MoveNext
        Wend
        
        lsCadena = lsCadena & ImprimeAsientoConsolidadoDet(sqlAdd, lsNros, psMoneda, pdFecha, "", psComentario & " - " & Rs!cCtaContCod, lsCtaNoAfecta) & oImpresora.gPrnSaltoPagina
        
        Rs.MoveNext
    Wend
    
    ImprimeAsientoConsolidadoPalcial = lsCadena
End Function

Public Function ImprimeAsientoConsolidadoDet(psCtaContCod As String, psMovNro As String, psMoneda As String, pdFecha As Date, Optional psCtaNoConsideradas As String = "", Optional psComentario As String = "", Optional CtasNoConsideradas As String = "") As String
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lnItem As Long
    Dim lsCadena As String
    Dim lsCabecera As String
    
    Dim CON As String
    Dim COFF As String
    Dim BON As String
    Dim BOFF As String
    
    Dim nTot As Currency
    Dim nTotH As Currency
    
    Dim oDGeneral As DGeneral
    Set oDGeneral = New DGeneral
    
    CON = PrnSet("C+")
    COFF = PrnSet("C-")
    BON = PrnSet("B+")
    BOFF = PrnSet("B-")
        
    Sql = " Select A.cCtaContCod , B.cCtaContdesc + ' ' + C.cCtaContDesc, A.nMovImporte From" _
        & " (" _
        & "     Select cCtaContCod, Sum(nMovImporte) nMovImporte From MovCta MC" _
        & "         Where nMovNro in ('" & psMovNro & "') Group By cCtaContCod Having cCtaContCod Not In ('" & CtasNoConsideradas & "') " _
        & " ) As A" _
        & " Join Ctacont C On A.cCtaContCod = C.cCtaContCod" _
        & " Join CtaCont B On B.cCtaContCod = LEFT(A.cCtaContCod,LEN(A.cCtaContCod)-2)" _
        & " Where LEFT(A.cCtaContCod,2) Not In ('" & psCtaNoConsideradas & "')" _
        & " And (A.cCtaContCod in ('" & psCtaContCod & "'))  Or Left(A.cCtaContCod,Len(A.cCtaContCod)-2) in (Select Left(cCtaContCodD,Len(cCtaContCodD)-2) Cta From OpeCtaCta Where cOpeCod = '571200' And cCtaContCodH In ('" & psCtaContCod & "'))" _
        & " Or Left(A.cCtaContCod,Len(A.cCtaContCod)-2) in (Select Left(cCtaContCodD,Len(cCtaContCodD)-2) Cta From OpeCtaCta Where cOpeCod = '571200' And cCtaContCodH in ('" & psCtaContCod & "')) " _
        & " Order By A.cCtaContCod"
    
    If psCtaContCod = "19110707" Then
        Sql = " Select A.cCtaContCod , B.cCtaContdesc + ' ' + C.cCtaContDesc, A.nMovImporte From" _
            & " (" _
            & "     Select cCtaContCod, Sum(nMovImporte) nMovImporte From MovCta MC" _
            & "         Where nMovNro in ('" & psMovNro & "') Group By cCtaContCod" _
            & " ) As A" _
            & " Join Ctacont C On A.cCtaContCod = C.cCtaContCod" _
            & " Join CtaCont B On B.cCtaContCod = LEFT(A.cCtaContCod,LEN(A.cCtaContCod)-2)" _
            & " Where LEFT(A.cCtaContCod,2) Not In ('" & psCtaNoConsideradas & "')" _
            & " And (A.cCtaContCod = '" & psCtaContCod & "' Or Left(A.cCtaContCod,Len(A.cCtaContCod)-2) in (Select Left(cCtaContCodD,Len(cCtaContCodD)-2) Cta From OpeCtaCta Where cOpeCod = '571200' And cCtaContCodH = '" & psCtaContCod & "')" _
            & " Or Left(A.cCtaContCod,Len(A.cCtaContCod)-2) in (Select Left(cCtaContCodD,Len(cCtaContCodD)-2) Cta From OpeCtaCta Where cOpeCod = '571200' And cCtaContCodH like '" & psCtaContCod & "%')) " _
            & " Order By A.cCtaContCod"
    End If
    
    oCon.AbreConexion
    '
    
    lsCadena = " A S I E N T O  C O N T A B L E"
    lsCadena = lsCadena & ImpreCabAsiento(80, pdFecha, "", "__" & psMoneda, "", lsCadena, True) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & JustificaTexto(psComentario, 80, 5) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea

    lsCabecera = ""
    lsCabecera = lsCabecera & CON
    lsCabecera = lsCabecera & "====================================================================================================================================" & oImpresora.gPrnSaltoLinea
    lsCabecera = lsCabecera & "Itm C U E N T A   C O N T A B L E                                                                           DEBE            HABER" & oImpresora.gPrnSaltoLinea
    lsCabecera = lsCabecera & "    Codigo              Descripcion" & oImpresora.gPrnSaltoLinea
    lsCabecera = lsCabecera & "------------------------------------------------------------------------------------------------------------------------------------" & COFF & oImpresora.gPrnSaltoLinea
    
    lsCadena = lsCadena & lsCabecera
    
    Set Rs = oCon.CargaRecordSet(Sql)
    lnItem = 0
    
    While Not Rs.EOF
       lnItem = lnItem + 1
       lsCadena = lsCadena & CON & Format(lnItem, "000") & " " & Justifica(Rs!cCtaContCod, 20) & " " _
                & Justifica(oDGeneral.CuentaNombre(Rs!cCtaContCod), 71) & " " _
                & IIf(Rs!NMOVIMPORTE > 0, PrnVal(Rs!NMOVIMPORTE, 16, 2), Space(16)) & " " _
                & IIf(Rs!NMOVIMPORTE < 0, PrnVal(Rs!NMOVIMPORTE * -1, 16, 2), Space(16)) _
                & COFF & oImpresora.gPrnSaltoLinea
                
        If Rs!NMOVIMPORTE > 0 Then
           nTot = nTot + Val(Rs!NMOVIMPORTE)
        Else
           nTotH = nTotH + Val(Rs!NMOVIMPORTE) * -1
        End If
                
        If lnItem Mod 50 = 0 Then
            lsCadena = lsCadena & CON & "====================================================================================================================================" & COFF & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
            lsCadena = lsCadena & lsCabecera
        End If
        
        Rs.MoveNext
    Wend
    
    If nTot > 0 Or nTotH > 0 Then
       lsCadena = lsCadena & CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
       lsCadena = lsCadena & BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF + oImpresora.gPrnSaltoLinea
    End If
    If Rs.State = adStateOpen Then Rs.Close: Set Rs = Nothing
    lsCadena = lsCadena & CON & "====================================================================================================================================" & COFF & oImpresora.gPrnSaltoLinea
    
    lsCadena = lsCadena & ImprePiePag(80)
    
    lsCadena = lsCadena & "   " & oImpresora.gPrnSaltoLinea

    ImprimeAsientoConsolidadoDet = lsCadena
End Function

Public Function ImprimeAsientoContableConsolidado(ByVal psMovNro As String, ByVal pnLinPage As Integer, ByVal pnColPage As Integer, Optional ByVal psTit1 As String, Optional ByVal psCabe1 As String, Optional ByVal psPiePag As String = "9", Optional ByVal psOtraFirma As String = "", Optional ByVal pbMoneda As Boolean = True) As String
    'On Error GoTo ImprimeAsientoContableErr
    Dim sSql As String
    Dim Rs       As New ADODB.Recordset
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
    
    Dim lsOrigenCab As String
    Dim lsDestinoCab As String
    
    
    Dim nMovItem As Integer
    Dim psMovNro1 As String
    Dim psMovNro2 As String
    
    Set oConect = New DConecta
    
    Dim oMov As DMov
    Set oMov = New DMov
    
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
    psMovNro1 = Left(psMovNro, InStr(1, psMovNro, ",") - 1)
    psMovNro2 = Right(psMovNro, InStr(1, psMovNro, ",") - 1)
    
    psMovNro1 = oMov.GetcMovNro(psMovNro1)
    psMovNro2 = oMov.GetcMovNro(psMovNro2)
    '****Modificado por ALPA
    '****04/03/2008
    lsFecha = CDate(Mid(psMovNro1, 7, 2) + "/" + Mid(psMovNro1, 5, 2) + "/" + Mid(psMovNro1, 1, 4))
    '****end
        
        
        lnCambioFijo = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCFijoMes)
        lnCambioVenta = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCVenta)
        lnCambioCompra = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCCompra)
    
       sSql = "SELECT * FROM Mov WHERE nMovNro in (" & psMovNro & ")"
       Set Rs = oConect.CargaRecordSet(sSql)
       If Rs.EOF Then
          Exit Function
       End If
       If InStr(1, Rs!cMovDesc, "-") = 0 Then
        lsGlosa = Rs!cMovDesc
        Else
        lsGlosa = Replace(Left(Rs!cMovDesc, InStr(1, Rs!cMovDesc, "-") - 1), Chr(13) & oImpresora.gPrnSaltoLinea, " ")
       End If
       
       
       lsMovNro = psMovNro
       lsOpeCod = Rs!COPECOD
       If Rs!nMovFlag = gMovFlagDeExtorno Then
            psTit1 = " A S I E N T O  C O N T A B L E   D E   E X T O R N O"
       End If
       sAsiento = sAsiento & ImpreCabAsiento(pnColPage, CDate(lsFecha), "", lsOpeCod, lsMovNro, psTit1, pbMoneda) & oImpresora.gPrnSaltoLinea
       nLi = 5
       
       'Tipo de Cambio
       
        sSql = "SELECT nMovTpoCambio FROM MovTpoCambio MC JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE M.cMovNro = '" & lsMovNro & "'"
        Set Rs = oConect.CargaRecordSet(sSql)
        If Not Rs.EOF Then
            sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "      T.C.Mercado: " & Format(Rs!nMovTpoCambio, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
        Else
            If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Then
                sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "     T.C.Compra : " & Format(lnCambioCompra, "##,###,##0.000") & "     T.C.Venta : " & Format(lnCambioVenta, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
            End If
        End If
       nLi = nLi + 1
       'Datos de Documentos
       sSql = "SELECT b.cDocAbrev, a.cDocNro, a.dDocFecha FROM MovDoc a join mov m on m.nmovnro =a.nmovnro  , " & vsBaseComunes & "Documento b " _
            & "WHERE m.cMovNro = '" & lsMovNro & "' and b.nDocTpo = a.nDocTpo "
       
        Set rsDoc = oConect.CargaRecordSet(sSql)
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
          sAsiento = sAsiento + ImpreFormat("Documentos", 25) & " :" & sDoc & oImpresora.gPrnSaltoLinea
        End If
        rsDoc.Close
       'Glosa
       sAsiento = sAsiento & ImpreGlosa(lsGlosa, pnColPage) + oImpresora.gPrnSaltoLinea
       nLi = nLi + 2
       lsMovBilletaje = ""
       lsMovBilletaje = ImprimeBilletaje(Mid(lsOpeCod, 3, 1), lsMovNro, pnColPage)
       If lsMovBilletaje <> "" Then
            sAsiento = sAsiento + lsMovBilletaje & oImpresora.gPrnSaltoLinea
            nLi = nLi + 8
       End If
    'Cabecera 2 de Asiento
    sCabecera = sCabecera + CON
    sCabecera = sCabecera + "====================================================================================================================================" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "Itm C U E N T A   C O N T A B L E                                                                           DEBE            HABER" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "    Codigo              Descripcion" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "------------------------------------------------------------------------------------------------------------------------------------" & COFF & oImpresora.gPrnSaltoLinea
    'MovCta para MN
    'sSQL = "SELECT mc.nMovItem, mc.cCtaContCod, mc.nMovImporte " _
    '     & "FROM   MovCta mc JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE  M.cMovNro = '" & lsMovNro & "' "
         
    sSql = "SELECT  mc.cCtaContCod,sum(mc.nMovImporte)as NMOVIMPORTE FROM   " _
            & " MovCta mc JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE " _
            & " M.nMovNro in(" & psMovNro & ") " _
            & " group by mc.cCtaContCod order by cCtaContCod desc "

         
         
    Set Rs = oConect.CargaRecordSet(sSql)
    If Rs.EOF Then
       'MsgBox "No existen datos de Movimiento ", vbInformation, "Error"
       Exit Function
    End If
    nTot = 0
    nTotH = 0
    sAsiento = sAsiento + sCabecera
    nLi = nLi + 4
    sAsiento = sAsiento + BON & " EN MONEDA NACIONAL " + oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento + "--------------------" & BOFF + oImpresora.gPrnSaltoLinea
    nMovItem = 1
    Do While Not Rs.EOF
      'DoEvents
       If nLi > pnLinPage - 6 Then
           nLi = 5
           sAsiento = sAsiento + oImpresora.gPrnSaltoPagina & sCabecera
       End If
       
       sAsiento = sAsiento + CON & Format(nMovItem, "000") & " " & Justifica(Rs!cCtaContCod, 20) & " " & Justifica(oDGeneral.CuentaNombre(Rs!cCtaContCod), 71) & " " _
            & Right(Space(16) & IIf(Rs!NMOVIMPORTE > 0, Format(Rs!NMOVIMPORTE, "##,###,##0.00"), ""), 16) & " " _
            & Right(Space(16) & IIf(Rs!NMOVIMPORTE < 0, Format(Rs!NMOVIMPORTE * -1, "##,###,##0.00"), ""), 16) _
            & COFF & oImpresora.gPrnSaltoLinea
       If Rs!NMOVIMPORTE > 0 Then
          nTot = nTot + Val(Rs!NMOVIMPORTE)
       Else
          nTotH = nTotH + Val(Rs!NMOVIMPORTE) * -1
       End If
       nLi = nLi + 1
       nMovItem = nMovItem + 1
       Rs.MoveNext
    Loop
    sAsiento = sAsiento + CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF & oImpresora.gPrnSaltoLinea
    nLi = nLi + 2
    
    sSql = "SELECT mc.nMovItem, mc.cCtaContCod, me.nmovMEimporte " _
         & "FROM   mov m join MovCta mc on mc.nmovnro=m.nmovnro JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
         & "WHERE  m.cMovNro = '" & lsMovNro & "'"
    Set Rs = oConect.CargaRecordSet(sSql)
    sTexto = ""
    nTot = 0
    nTotH = 0
    If Not Rs.EOF Then
       sAsiento = sAsiento + BON & " EN MONEDA EXTRANJERA " & oImpresora.gPrnSaltoLinea
       sAsiento = sAsiento + "----------------------" & BOFF & oImpresora.gPrnSaltoLinea
       nLi = nLi + 2
    End If
    Do While Not Rs.EOF
      'DoEvents
       If nLi > pnLinPage - 6 Then
          nLi = 5
          sTexto = sTexto + oImpresora.gPrnSaltoPagina
          sTexto = sTexto + sCabecera
       End If
       sTexto = sTexto + CON & Format(Rs!nMovItem, , "000") & " " & Justifica(Rs!cCtaContCod, 20) & " " _
              & Justifica(oDGeneral.CuentaNombre(Rs!cCtaContCod), 71) & " " _
              & IIf(Rs!NMOVMEIMPORTE > 0, PrnVal(Rs!NMOVMEIMPORTE, 16, 2), Space(16)) & " " _
              & IIf(Rs!NMOVMEIMPORTE < 0, PrnVal(Rs!NMOVMEIMPORTE * -1, 16, 2), Space(16)) _
              & COFF & oImpresora.gPrnSaltoLinea
       If Rs!NMOVMEIMPORTE > 0 Then
          nTot = nTot + Val(Rs!NMOVMEIMPORTE)
       Else
          nTotH = nTotH + Val(Rs!NMOVMEIMPORTE) * -1
       End If
       nLi = nLi + 1
       Rs.MoveNext
    Loop
    If nTot > 0 Or nTotH > 0 Then
       sAsiento = sAsiento & sTexto
       sAsiento = sAsiento + CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
       sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF + oImpresora.gPrnSaltoLinea
       nLi = nLi + 3
    End If
    If Rs.State = adStateOpen Then Rs.Close: Set Rs = Nothing
    sAsiento = sAsiento + CON & "====================================================================================================================================" & COFF & oImpresora.gPrnSaltoLinea
    
    sAsiento = sAsiento + ImprePiePag(pnColPage, psPiePag, psOtraFirma)
    
    sAsiento = sAsiento + "   " + oImpresora.gPrnSaltoLinea

    Set oDGeneral = Nothing
    Set oConect = Nothing
    ImprimeAsientoContableConsolidado = sAsiento
    
    Exit Function
ImprimeAsientoContableErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeAsientoContable Method")
End Function


Public Function ImprimeAsientoContConsolPorAgencia(ByVal psMovNro As String, ByVal pnLinPage As Integer, ByVal pnColPage As Integer, Optional ByVal psTit1 As String, Optional ByVal psCabe1 As String, Optional ByVal psPiePag As String = "9", Optional ByVal psOtraFirma As String = "", Optional ByVal pbMoneda As Boolean = True, Optional ByVal psPeriodo As String = "", Optional ByVal psAgencia As String = "", Optional ByVal psPlanilla As String = "", Optional ByVal psOperacion As String = "", Optional ByVal psDescripcion As String = "") As String
    'On Error GoTo ImprimeAsientoContableErr
    Dim sSql As String
    Dim Rs       As New ADODB.Recordset
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
    
    Dim lsOrigenCab As String
    Dim lsDestinoCab As String
    
    
    Dim nMovItem As Integer
    Dim psMovNro1 As String
    Dim psMovNro2 As String
    
    Set oConect = New DConecta
    
    Dim oMov As DMov
    Set oMov = New DMov
    
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
    psMovNro1 = Left(psMovNro, InStr(1, psMovNro, ",") - 1)
    psMovNro2 = Right(psMovNro, InStr(1, psMovNro, ",") - 1)
    
    psMovNro1 = oMov.GetcMovNro(psMovNro1)
    psMovNro2 = oMov.GetcMovNro(psMovNro2)
        
    lsFecha = CDate(Mid(psMovNro1, 7, 2) + "/" + Mid(psMovNro1, 5, 2) + "/" + Mid(psMovNro, 1, 4))
        
        
        lnCambioFijo = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCFijoMes)
        lnCambioVenta = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCVenta)
        lnCambioCompra = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCCompra)
    
       sSql = "SELECT * FROM Mov WHERE nMovNro in (" & psMovNro & ")"
       Set Rs = oConect.CargaRecordSet(sSql)
       If Rs.EOF Then
          Exit Function
       End If
       'psDescripcion
       lsGlosa = Replace(Left(Rs!cMovDesc, InStr(1, Rs!cMovDesc, "-") - 1), Chr(13) & oImpresora.gPrnSaltoLinea, " ")
       lsGlosa = lsGlosa + "  " + psDescripcion
       lsMovNro = psMovNro
       lsOpeCod = Rs!COPECOD
       If Rs!nMovFlag = gMovFlagDeExtorno Then
            psTit1 = " A S I E N T O  C O N T A B L E   D E   E X T O R N O"
       End If
       sAsiento = sAsiento & ImpreCabAsiento(pnColPage, CDate(lsFecha), "", lsOpeCod, lsMovNro, psTit1, pbMoneda) & oImpresora.gPrnSaltoLinea
       nLi = 5
       
       'Tipo de Cambio
       
        sSql = "SELECT nMovTpoCambio FROM MovTpoCambio MC JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE M.cMovNro = '" & lsMovNro & "'"
        Set Rs = oConect.CargaRecordSet(sSql)
        If Not Rs.EOF Then
            sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "      T.C.Mercado: " & Format(Rs!nMovTpoCambio, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
        Else
            If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Then
                sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "     T.C.Compra : " & Format(lnCambioCompra, "##,###,##0.000") & "     T.C.Venta : " & Format(lnCambioVenta, "##,###,##0.000") & oImpresora.gPrnSaltoLinea
            End If
        End If
       nLi = nLi + 1
       'Datos de Documentos
       sSql = "SELECT b.cDocAbrev, a.cDocNro, a.dDocFecha FROM MovDoc a join mov m on m.nmovnro =a.nmovnro  , " & vsBaseComunes & "Documento b " _
            & "WHERE m.cMovNro = '" & lsMovNro & "' and b.nDocTpo = a.nDocTpo "
       
        Set rsDoc = oConect.CargaRecordSet(sSql)
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
          sAsiento = sAsiento + ImpreFormat("Documentos", 25) & " :" & sDoc & oImpresora.gPrnSaltoLinea
        End If
        rsDoc.Close
       'Glosa
       sAsiento = sAsiento & ImpreGlosa(lsGlosa, pnColPage) + oImpresora.gPrnSaltoLinea
       nLi = nLi + 2
       lsMovBilletaje = ""
       lsMovBilletaje = ImprimeBilletaje(Mid(lsOpeCod, 3, 1), lsMovNro, pnColPage)
       If lsMovBilletaje <> "" Then
            sAsiento = sAsiento + lsMovBilletaje & oImpresora.gPrnSaltoLinea
            nLi = nLi + 8
       End If
    'Cabecera 2 de Asiento
    sCabecera = sCabecera + CON
    sCabecera = sCabecera + "====================================================================================================================================" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "Itm C U E N T A   C O N T A B L E                                                                           DEBE            HABER" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "    Codigo              Descripcion" & oImpresora.gPrnSaltoLinea
    sCabecera = sCabecera + "------------------------------------------------------------------------------------------------------------------------------------" & COFF & oImpresora.gPrnSaltoLinea
    'MovCta para MN
    'sSQL = "SELECT mc.nMovItem, mc.cCtaContCod, mc.nMovImporte " _
    '     & "FROM   MovCta mc JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE  M.cMovNro = '" & lsMovNro & "' "
         
    'psperiodo,psagencia,psplanilla,psOperacion
    
    'sSql = "SELECT  mc.cCtaContCod,sum(mc.nMovImporte)as NMOVIMPORTE FROM   " _
            & " MovCta mc JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE " _
            & " M.nMovNro in(" & psMovNro & ") " _
            & " group by mc.cCtaContCod order by cCtaContCod desc "

    'spRHAsientoAgencia     @Periodo varchar(8),    @Planilla char(3),     @Agencia char(2),     @cOpeCod varchar(6)
    sSql = "spRHAsientoAgencia '" & psPeriodo & "','" & psPlanilla & "','" & psAgencia & "', '" & psOperacion & "' "
    
         
    Set Rs = oConect.CargaRecordSet(sSql)
    If Rs.EOF Then
       'MsgBox "No existen datos de Movimiento ", vbInformation, "Error"
       Exit Function
    End If
    nTot = 0
    nTotH = 0
    sAsiento = sAsiento + sCabecera
    nLi = nLi + 4
    sAsiento = sAsiento + BON & " EN MONEDA NACIONAL " + oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento + "--------------------" & BOFF + oImpresora.gPrnSaltoLinea
    nMovItem = 1
    Do While Not Rs.EOF
      'DoEvents
       If nLi > pnLinPage - 6 Then
           nLi = 5
           sAsiento = sAsiento + oImpresora.gPrnSaltoPagina & sCabecera
       End If
       
       sAsiento = sAsiento + CON & Format(nMovItem, "000") & " " & Justifica(Rs!cCtaContCod, 20) & " " & Justifica(oDGeneral.CuentaNombre(Rs!cCtaContCod), 71) & " " _
            & Right(Space(16) & IIf(Rs!NMOVIMPORTE > 0, Format(Rs!NMOVIMPORTE, "##,###,##0.00"), ""), 16) & " " _
            & Right(Space(16) & IIf(Rs!NMOVIMPORTE < 0, Format(Rs!NMOVIMPORTE * -1, "##,###,##0.00"), ""), 16) _
            & COFF & oImpresora.gPrnSaltoLinea
       If Rs!NMOVIMPORTE > 0 Then
          nTot = nTot + Val(Rs!NMOVIMPORTE)
       Else
          nTotH = nTotH + Val(Rs!NMOVIMPORTE) * -1
       End If
       nLi = nLi + 1
       nMovItem = nMovItem + 1
       Rs.MoveNext
    Loop
    sAsiento = sAsiento + CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF & oImpresora.gPrnSaltoLinea
    nLi = nLi + 2
    
    sSql = "SELECT mc.nMovItem, mc.cCtaContCod, me.nmovMEimporte " _
         & "FROM   mov m join MovCta mc on mc.nmovnro=m.nmovnro JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
         & "WHERE  m.cMovNro = '" & lsMovNro & "'"
    Set Rs = oConect.CargaRecordSet(sSql)
    sTexto = ""
    nTot = 0
    nTotH = 0
    If Not Rs.EOF Then
       sAsiento = sAsiento + BON & " EN MONEDA EXTRANJERA " & oImpresora.gPrnSaltoLinea
       sAsiento = sAsiento + "----------------------" & BOFF & oImpresora.gPrnSaltoLinea
       nLi = nLi + 2
    End If
    Do While Not Rs.EOF
      'DoEvents
       If nLi > pnLinPage - 6 Then
          nLi = 5
          sTexto = sTexto + oImpresora.gPrnSaltoPagina
          sTexto = sTexto + sCabecera
       End If
       sTexto = sTexto + CON & Format(Rs!nMovItem, , "000") & " " & Justifica(Rs!cCtaContCod, 20) & " " _
              & Justifica(oDGeneral.CuentaNombre(Rs!cCtaContCod), 71) & " " _
              & IIf(Rs!NMOVMEIMPORTE > 0, PrnVal(Rs!NMOVMEIMPORTE, 16, 2), Space(16)) & " " _
              & IIf(Rs!NMOVMEIMPORTE < 0, PrnVal(Rs!NMOVMEIMPORTE * -1, 16, 2), Space(16)) _
              & COFF & oImpresora.gPrnSaltoLinea
       If Rs!NMOVMEIMPORTE > 0 Then
          nTot = nTot + Val(Rs!NMOVMEIMPORTE)
       Else
          nTotH = nTotH + Val(Rs!NMOVMEIMPORTE) * -1
       End If
       nLi = nLi + 1
       Rs.MoveNext
    Loop
    If nTot > 0 Or nTotH > 0 Then
       sAsiento = sAsiento & sTexto
       sAsiento = sAsiento + CON & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
       sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(nTotH, 16, 2) & BOFF & COFF + oImpresora.gPrnSaltoLinea
       nLi = nLi + 3
    End If
    If Rs.State = adStateOpen Then Rs.Close: Set Rs = Nothing
    sAsiento = sAsiento + CON & "====================================================================================================================================" & COFF & oImpresora.gPrnSaltoLinea
    
    sAsiento = sAsiento + ImprePiePag(pnColPage, psPiePag, psOtraFirma)
        
    'sAsiento = sAsiento + "   " + oImpresora.gPrnSaltoLinea
    sAsiento = sAsiento + "   "

    Set oDGeneral = Nothing
    Set oConect = Nothing
    ImprimeAsientoContConsolPorAgencia = sAsiento
    
    Exit Function
ImprimeAsientoContableErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeAsientoContable Method")
End Function



