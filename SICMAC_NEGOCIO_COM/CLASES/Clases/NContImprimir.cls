VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NContImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A7EE84101F4"
Option Base 0
Option Explicit
Dim vsConexion As String
Dim vsBaseComunes As String
Dim vsBasePesonas As String
Dim vsAgeNom As String
Dim vsEmpNom As String
Dim vsFecSis As String

Type tBilletajes
    lsBilletes As Currency
    lnMontoBil As Currency
    lsMonedas As Currency
    lnMontoMon As Currency
End Type

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
Public Function ImprimeAsientoContable(ByVal psMovNro As String, ByVal pnLinPage As Integer, ByVal pnColPage As Integer, Optional ByVal psTit1 As String, Optional ByVal psCabe1 As String, Optional ByVal psPiePag As String = "9", Optional ByVal psOtraFirma As String = "", Optional psNomCmact As String = "") As String
    On Error GoTo ImprimeAsientoContableErr
    Dim sSql As String
    Dim rs       As New ADODB.Recordset
    Dim rsDoc       As New ADODB.Recordset
    Dim nTot     As Currency
    Dim lnTotH    As Currency
    Dim sDoc     As String
    Dim sTexto   As String
    Dim sAsiento As String
    Dim nLi      As Integer
    Dim sCabecera As String
    Dim lnCambioFijo As Currency
    Dim lnCambioVenta As Currency
    Dim lnCambioCompra As Currency
    Dim Con As String
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
    Con = PrnSet("C+")
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
       Set rs = oConect.CargaRecordSet(sSql)
       If rs.EOF Then
          Exit Function
       End If
       lsGlosa = Replace(rs!cMovDesc, Chr(13) & Chr(10), " ")
       lsMovNro = psMovNro
       lsOpeCod = rs!copecod
       If rs!nMovFlag = gMovFlagDeExtorno Then
            psTit1 = " A S I E N T O  C O N T A B L E   D E   E X T O R N O"
       End If
       sAsiento = sAsiento & ImpreCabAsiento(pnColPage, CDate(lsFecha), psNomCmact, lsOpeCod, lsMovNro, psTit1, True) & Chr$(10)
       nLi = 5
       If psCabe1 <> "" Then
          sAsiento = sAsiento & psCabe1 & Chr$(10)
          nLi = nLi + 4
       Else
           'BUSCA LAS PERSONA INVOLUCRADAS EN EL MOVIMIENTO
            sSql = " SELECT  M.CMOVNRO,  " _
                    & "         CASE    WHEN mc.nMovImporte>0 THEN 'D' " _
                    & "                 WHEN mc.nMovImporte<0 THEN 'H' END as cDH , " _
                    & "         P.cPersCod, " _
                    & "         Convert(Char(40),P.CPERSNOMBRE) as cPersNombre, " _
                    & "         MO.COBJETOCOD , MC.nMovImporte " _
                    & " FROM    MOVCTA MC JOIN MOV M    ON M.NMOVNRO = MC.NMOVNRO " _
                    & "         JOIN MOVOBJ MO          ON MO.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
                    & "         JOIN MOVOBJPERS MOI         ON MOI.NMOVNRO = MO.NMOVNRO AND MOI.NMOVITEM = MO.NMOVITEM " _
                    & "         JOIN PERSONA P          ON P.CPERSCOD = MOI.CPERSCOD " _
                    & " WHERE   M.cMovNro = '" & lsMovNro & "' AND MO.cObjetoCod ='" & Format(ObjPersona, "00") & "'"
              
            Set rs = oConect.CargaRecordSet(sSql)
            sDestino = "": sOrigen = ""
            Do While Not rs.EOF
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, Chr(10) & Space(14), "") & PstaNombre(rs!cPersNombre)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, Chr(10) & Space(14), "") & PstaNombre(rs!cPersNombre)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat("Persona(s) Origen  ", 25) & " :" & sOrigen & Chr$(10)
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat("Persona(s) Destino ", 25) & " :" & sDestino & Chr$(10)
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
            
            Set rs = oConect.CargaRecordSet(sSql)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, Chr(10) & Space(14), "") & Trim(rs!cPersNombre) & " " & Trim(rs!DescCtaIF)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, Chr(10) & Space(14), "") & Trim(rs!cPersNombre) & " " & Trim(rs!DescCtaIF)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat("Cuenta(s) Origen", 20) & " :" + PrnSet("C+") + sOrigen + PrnSet("C-") + Chr$(10)
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat("Cuenta(s) Destino", 20) & " :" + PrnSet("C+") + sDestino + PrnSet("C-") + Chr$(10)
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

            Set rs = oConect.CargaRecordSet(sSql)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                lsDesObjeto = Trim(rs!cObjetoDesc)
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, Chr(10) & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!CAREADESCRIPCION)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, Chr(10) & Space(14), "") & Trim(rs!cAreaCod) & " - " & Trim(rs!CAREADESCRIPCION)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Origen ", 20) & " :" & sOrigen & Chr$(10)
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Destino ", 20) & " :" & sDestino & Chr$(10)
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
            '        sDestino = sDestino & IIf(Len(sDestino) > 0, Chr(10) & Space(14), "") & Trim(rs!CAGECOD) & " - " & Trim(rs!CAGEDESCRIPCION)
            '    Else
            '        sOrigen = sOrigen & IIf(Len(sOrigen) > 0, Chr(10) & Space(14), "") & Trim(rs!CAGECOD) & " - " & Trim(rs!CAGEDESCRIPCION)
            '    End If
            '    nLi = nLi + 1
            '    rs.MoveNext
            'Loop
            'If sOrigen <> "" Then
            '    sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Origen ", 25) & " :" & sOrigen & Chr$(10)
            'End If
            'If sDestino <> "" Then
            '    sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Destino ", 25) & " :" & sDestino & Chr$(10)
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
            
            Set rs = oConect.CargaRecordSet(sSql)
            sDestino = "": sOrigen = "": lsDesObjeto = ""
            Do While Not rs.EOF
                lsDesObjeto = "OBJETOS"
                If rs!cDH = "D" Then
                    sDestino = sDestino & IIf(Len(sDestino) > 0, Chr(10) & Space(14), "") & Trim(rs!cObjetoCod) & "-" & Trim(rs!cObjetoDesc)
                Else
                    sOrigen = sOrigen & IIf(Len(sOrigen) > 0, Chr(10) & Space(14), "") & Trim(rs!cObjetoCod) & "-" & Trim(rs!cObjetoDesc)
                End If
                nLi = nLi + 1
                rs.MoveNext
            Loop
            If sOrigen <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Origen ", 20) & " :" & sOrigen & Chr$(10)
            End If
            If sDestino <> "" Then
                sAsiento = sAsiento & ImpreFormat(lsDesObjeto & " Destino", 20) & " :" & sDestino & Chr$(10)
            End If
            
        End If ' del pscab1
    
       'Tipo de Cambio
       
        sSql = "SELECT nMovTpoCambio FROM MovTpoCambio MC JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE M.cMovNro = '" & lsMovNro & "'"
        Set rs = oConect.CargaRecordSet(sSql)
        If Not rs.EOF Then
            sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "      T.C.Mercado: " & Format(rs!nMovTpoCambio, "##,###,##0.000") & Chr(10)
        Else
            If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Then
                sAsiento = sAsiento + "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "     T.C.Compra : " & Format(lnCambioCompra, "##,###,##0.000") & "     T.C.Venta : " & Format(lnCambioVenta, "##,###,##0.000") & Chr(10)
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
                sDoc = sDoc + Chr$(10)
                nLi = nLi + 1
             End If
             rsDoc.MoveNext
          Loop
          sAsiento = sAsiento + ImpreFormat("Documentos", 20) & " :" + PrnSet("C+") + sDoc + PrnSet("C-") + Chr$(10)
        End If
        rsDoc.Close
       'Glosa
       sAsiento = sAsiento & ImpreGlosa(lsGlosa, pnColPage, ImpreFormat("Glosa       ", 20) & " :") + Chr$(10)
       nLi = nLi + 2
       lsMovBilletaje = ""
       lsMovBilletaje = ImprimeBilletaje(Mid(lsOpeCod, 3, 1), lsMovNro, pnColPage)
       If lsMovBilletaje <> "" Then
            sAsiento = sAsiento + lsMovBilletaje & Chr$(10)
            nLi = nLi + 8
       End If
    'Cabecera 2 de Asiento
    sCabecera = sCabecera + Con
    sCabecera = sCabecera + "====================================================================================================================================" & Chr$(10)
    sCabecera = sCabecera + "Itm C U E N T A   C O N T A B L E                                                                           DEBE            HABER" & Chr$(10)
    sCabecera = sCabecera + "    Codigo              Descripcion" & Chr$(10)
    sCabecera = sCabecera + "------------------------------------------------------------------------------------------------------------------------------------" & COFF & Chr$(10)
    'MovCta para MN
    sSql = "SELECT mc.nMovItem, mc.cCtaContCod, mc.nMovImporte " _
         & "FROM   MovCta mc JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE  M.cMovNro = '" & lsMovNro & "' "
    Set rs = oConect.CargaRecordSet(sSql)
    If rs.EOF Then
       'MsgBox "No existen datos de Movimiento ", vbInformation, "Error"
       Exit Function
    End If
    nTot = 0
    lnTotH = 0
    sAsiento = sAsiento + sCabecera
    nLi = nLi + 4
    sAsiento = sAsiento + BON & " EN MONEDA NACIONAL " + Chr$(10)
    sAsiento = sAsiento + "--------------------" & BOFF + Chr$(10)
    Do While Not rs.EOF
       DoEvents
       If nLi > pnLinPage - 6 Then
           nLi = 5
           sAsiento = sAsiento + Chr(12) & sCabecera
       End If
       sAsiento = sAsiento + Con & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 20) & " " & Justifica(oDGeneral.CuentaNombre(rs!cCtaContCod), 71) & " " _
            & Right(Space(16) & IIf(rs!nMovImporte > 0, Format(rs!nMovImporte, "##,###,##0.00"), ""), 16) & " " _
            & Right(Space(16) & IIf(rs!nMovImporte < 0, Format(rs!nMovImporte * -1, "##,###,##0.00"), ""), 16) _
            & COFF & Chr$(10)
       If rs!nMovImporte > 0 Then
          nTot = nTot + Val(rs!nMovImporte)
       Else
          lnTotH = lnTotH + Val(rs!nMovImporte) * -1
       End If
       nLi = nLi + 1
       rs.MoveNext
    Loop
    sAsiento = sAsiento + Con & "------------------------------------------------------------------------------------------------------------------------------------" & Chr$(10)
    sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(lnTotH, 16, 2) & BOFF & COFF & Chr$(10)
    nLi = nLi + 2
    
    sSql = "SELECT mc.nMovItem, mc.cCtaContCod, me.nmovMEimporte " _
         & "FROM   mov m join MovCta mc on mc.nmovnro=m.nmovnro JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
         & "WHERE  m.cMovNro = '" & lsMovNro & "'"
    Set rs = oConect.CargaRecordSet(sSql)
    sTexto = ""
    nTot = 0
    lnTotH = 0
    If Not rs.EOF Then
       sAsiento = sAsiento + BON & " EN MONEDA EXTRANJERA " & Chr$(10)
       sAsiento = sAsiento + "----------------------" & BOFF & Chr$(10)
       nLi = nLi + 2
    End If
    Do While Not rs.EOF
       DoEvents
       If nLi > pnLinPage - 6 Then
          nLi = 5
          sTexto = sTexto + Chr(12)
          sTexto = sTexto + sCabecera
       End If
       sTexto = sTexto + Con & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 20) & " " _
              & Justifica(oDGeneral.CuentaNombre(rs!cCtaContCod), 71) & " " _
              & IIf(rs!nMovMEImporte > 0, PrnVal(rs!nMovMEImporte, 16, 2), Space(16)) & " " _
              & IIf(rs!nMovMEImporte < 0, PrnVal(rs!nMovMEImporte * -1, 16, 2), Space(16)) _
              & COFF & Chr(10)
       If rs!nMovMEImporte > 0 Then
          nTot = nTot + Val(rs!nMovMEImporte)
       Else
          lnTotH = lnTotH + Val(rs!nMovMEImporte) * -1
       End If
       nLi = nLi + 1
       rs.MoveNext
    Loop
    If nTot > 0 Or lnTotH > 0 Then
       sAsiento = sAsiento & sTexto
       sAsiento = sAsiento + Con & "------------------------------------------------------------------------------------------------------------------------------------" & Chr$(10)
       sAsiento = sAsiento + BON & Space(97) & PrnVal(nTot, 16, 2) & " " & PrnVal(lnTotH, 16, 2) & BOFF & COFF + Chr$(10)
       nLi = nLi + 3
    End If
    If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
    sAsiento = sAsiento + Con & "====================================================================================================================================" & COFF & Chr$(10)
    
    sAsiento = sAsiento + ImprePiePag(pnColPage, psPiePag, psOtraFirma)
    sAsiento = sAsiento + "   " + Chr$(10)

    Set oDGeneral = Nothing
    Set oConect = Nothing
    ImprimeAsientoContable = sAsiento
    
    Exit Function
ImprimeAsientoContableErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeAsientoContable Method")
End Function

Public Function ImprimeReciboEgresos(ByVal pnColPage As Integer, ByVal psMovNro As String, ByVal psGlosa As String, _
                                ByVal pdFecha As String, ByVal psEmpresaLogo As String, ByVal psOpeCod As String, _
                                ByVal lbEsCajaChica As Boolean, ByVal psDescCajaChica As String, ByVal pnArendirFase As ARendirFases, _
                                ByVal psDocNroARendir As String, ByVal psDocNro As String, ByVal psFechaDoc As String, _
                                ByVal psPersCodProv As String, ByVal psPersNombreProv As String, _
                                ByVal psCargo As String, ByVal pnTotal As Currency) As String
    Dim sTexto As String
    Dim nAncho As Integer, N As Integer
    Dim lsTexto As String
    Dim BON As String
    Dim BOFF As String
    Dim Con As String
    Dim COFF As String
    On Error GoTo ImprimeReciboEgresosErr
    

    nAncho = pnColPage
    lsTexto = ""
    BON = PrnSet("B+")
    BOFF = PrnSet("B-")
    Con = PrnSet("C+")
    COFF = PrnSet("C-")
  
    lsTexto = lsTexto + ImpreCabAsiento(pnColPage, pdFecha, psEmpresaLogo, psOpeCod, psMovNro, " R E C I B O   D E   E G R E S O  Nro. " & psDocNro, , , , 1) & Chr$(10)
    If lbEsCajaChica Then
         lsTexto = lsTexto + "  CAJA CHICA   : " & BON & Mid(psDescCajaChica, 1, 60) & BOFF & Chr$(10)
    End If
    If pnArendirFase = ArendirRendicion Or pnArendirFase = ArendirSustentacion Then
        lsTexto = lsTexto + "  Rec. Arendir : " & BON & psDocNroARendir & BOFF & "  Fecha : " & BON & psFechaDoc & BOFF & Chr$(10)
    End If
    lsTexto = lsTexto + "  Persona      : " & BON & psPersCodProv & "  " & ImpreCarEsp(psPersNombreProv) & BOFF & Chr$(10)
    lsTexto = lsTexto + "  Cargo        : " & psCargo & Chr$(10)
    lsTexto = lsTexto + "  Importe      : " & BON & ConvNumLet(Trim(Str(pnTotal)), True, False) & BOFF & Chr$(10)
    lsTexto = lsTexto + ImpreGlosa(psGlosa, pnColPage, "  Concepto     : ") & Chr$(10)
    lsTexto = lsTexto + "" + Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10)
    lsTexto = lsTexto + BON & Space(15) & "_________________________                ______________________" & Chr$(10)
    lsTexto = lsTexto + Space(15) & Con & Centra(ImpreCarEsp(Mid(psPersNombreProv, 1, 40)), 40) & COFF & "                    Vo Bo JEFE AREA " & BOFF & Chr$(10)
  
    ImprimeReciboEgresos = lsTexto
    
    Exit Function
ImprimeReciboEgresosErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeReciboEgresos Method")
End Function

'##ModelId=3A969A8500EA
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
Dim Con As String
Dim COFF As String

BON = PrnSet("B+")
BOFF = PrnSet("B-")
Con = PrnSet("C+")
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

Set rs = oConect.CargaRecordSet(sql)
If rs.EOF Then Exit Function

If rs!nConsValor = gArendirTipoCajaChica Then
    sTexto = "RECIBO DE A RENDIR CUENTA DE CAJA CHICA  NRO. " & rs!cDocNro
Else
    sTexto = "R E C I B O   D E   A   R E N D I R   C U E N T A   NRO. " & rs!cDocNro
End If
    lsCadPrint = ImpreCabAsiento(pnColPage, rs!dDocFecha, psEmpresaLogo, rs!copecod, rs!cMovnro, sTexto, , , , 2)
    ' Definición de Cabecera 2
    If rs!nConsValor = gArendirTipoCajaChica Then
        lsCadPrint = lsCadPrint + "  CAJA CHICA   : " & BON & ImpreFormat(rs!cAreaCH & "-" & rs!nProcCh, 15) & ImpreCarEsp(rs!cDescCH) & BOFF & Chr(10)
    Else
        lsCadPrint = lsCadPrint + "  A Rendir de  : " & BON & ImpreFormat(rs!cAreaCod, 15) & ImpreCarEsp(rs!CAREADESCRIPCION) & BOFF & Chr$(10)
    End If
    lsCadPrint = lsCadPrint + "  Persona      : " & BON & ImpreFormat(rs!cPersCod, 15) & ImpreCarEsp(PstaNombre(rs!cPersNombre)) & BOFF & Chr$(10)
    lsCadPrint = lsCadPrint + "  Area/Agencia : " & BON & ImpreFormat(rs!cAreaCodPers, 15) & ImpreCarEsp(rs!cAreaDesPers) & BOFF & Chr$(10)
    lsCadPrint = lsCadPrint + "  Cargo        : " & Chr$(10)
    lsCadPrint = lsCadPrint + "  Importe      : " & BON & ConvNumLet(rs!copecod, rs!nMovSaldo, False) & BOFF & Chr$(10) & Chr$(10)
    If Mid(rs!copecod, 3, 1) = gMonedaExtranjera Then
        'lsCadPrint = lsCadPrint + "                 Tipo de Cambio : " & txtTipCambio & Chr$(10)
    End If
    lsCadPrint = lsCadPrint + ImpreGlosa(Trim(rs!cMovDesc), pnColPage, "  Concepto     : ")
    lsCadPrint = lsCadPrint + Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10)
    lsCadPrint = lsCadPrint + BON + Chr$(10)
    lsCadPrint = lsCadPrint + Space(20) & "___________________________" & Chr$(10)
    lsCadPrint = lsCadPrint + Space(20) & Con & Centra(Mid(PstaNombre(rs!cPersNombre), 1, 46), 46) & COFF & Chr$(10)
    lsCadPrint = lsCadPrint + Space(20) & Centra("  L.E. " & Trim(rs!DNI), 28) & BOFF & Chr$(10)
    lsCadPrint = lsCadPrint + Chr$(10) + Chr$(10) + Chr$(10)
    
    lsGlosa = "De acuerdo a la Política existente, la rendición del presente recibo se deberá hacer en un plazo no mayor de 48 horas, bajo responsabilidad." & Chr$(10) _
           & "Los Comprobantes de gastos serán a nombre de la " & BON & Trim(psEmpresa) & BOFF & ", RUC " & BON & psEmpresaRuc & BOFF & Chr$(10) & " "
    lsCadPrint = lsCadPrint + ImpreGlosa(lsGlosa, pnColPage, " NOTA : ") + Chr$(10) + Chr$(10) + Chr$(10) + Chr$(10)
    If rs!nConsValor <> gArendirTipoCajaChica Then
        lsCadPrint = lsCadPrint + Centra("____________________          _____________________", 80) + Chr$(10)
        lsCadPrint = lsCadPrint + Centra(" Vo Bo JEFE DE AREA              Vo Bo GERENCIA    ", 80) + Chr$(10)
    Else
        lsCadPrint = lsCadPrint + Centra("____________________", 80) + Chr$(10)
        lsCadPrint = lsCadPrint + Centra(" Vo Bo JEFE DE AREA ", 80) + Chr$(10)
    End If
    lsCadPrint = lsCadPrint + Chr$(10)
    ImprimeReciboARendir = lsCadPrint
End Function

Private Sub Class_Initialize()
Dim oIni As ClasIni
Set oIni = New ClasIni
vsConexion = oIni.CadenaConexion
vsBaseComunes = oIni.BaseComunes
vsBasePesonas = oIni.BasePersonas
sLpt = "LPT1"
BON = PrnSet("B+")
BOFF = PrnSet("B-")
Con = PrnSet("C+")
COFF = PrnSet("C-")
Set oIni = Nothing
End Sub
Private Function ImprimeBilletaje(ByVal psMoneda As Moneda, ByVal psMovNro As String, ByVal pnColPage As Integer, Optional ByVal pbCondensado As Boolean = True) As String
Dim sSql As String
Dim rs   As New ADODB.Recordset
Dim sTexto     As String
Dim sImpre As String
Dim BON As String
Dim BOFF As String
Dim COFF As String
Dim Con As String
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
Con = PrnSet("C+")
COFF = PrnSet("C-")

On Error GoTo ErrBillete

Set oEfec = New Defectivo
ImprimeBilletaje = ""

sTexto = "": sImpre = ""
Set rs = oEfec.GetEfectivoMovImp(psMovNro, True)
If Not rs.EOF And Not rs.BOF Then
    If rs!Moneda = gMonedaExtranjera Then
        lsSimbolo = "$."
    Else
        lsSimbolo = "S/."
    End If
    sTexto = sTexto + BON & PrnSet("N+") & "DESCOMPOSICION DE EFECTIVO : INGRESO" & BOFF & PrnSet("N-") & Chr$(10)
    sTexto = sTexto + IIf(pbCondensado, Con, "") + String(pnColPage - 3, "=") & Chr$(10)
    nMontoM = 0: nMontoB = 0
    nMontoTot = 0
    sTexto = sTexto + BON & ImpreFormat(" BILLETES ", 22, 7) & ImpreFormat("MONTO", 15) _
         & ImpreFormat(" MONEDAS ", 14, 5) & ImpreFormat("MONTO", 10) + BOFF & Chr$(10)
    sTexto = sTexto + String(pnColPage - 3, "-") + BOFF & Chr$(10)
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
        sTexto = sTexto + lsBilletes + lsMonedas & Chr$(10)
    Next
    sTexto = sTexto + String(pnColPage - 3, "-") & Chr$(10)
    sTexto = sTexto + BON & ImpreFormat("SUBTOTAL :", 12, 3) & ImpreFormat(lsSimbolo, 3, 1) & ImpreFormat(nMontoB, 15, 2, True) _
         & ImpreFormat("SUBTOTAL", 10, 8) & ImpreFormat(lsSimbolo, 3, 1) & ImpreFormat(nMontoM, 10, 2, True) + BOFF & Chr$(10)
    sTexto = sTexto + String(pnColPage - 3, "=") & Chr$(10)
    sTexto = sTexto + Space(45) & BON & "TOTAL :  " _
           & lsSimbolo & " " & PrnVal(nMontoTot, 14, 2) & BOFF & Chr$(10)
    sTexto = sTexto + Space(45) & String(31, "=") + COFF & Chr$(10)
    sTexto = sTexto + "" & Chr$(10)
End If
rs.Close
Set rs = Nothing

Set rs = oEfec.GetEfectivoMovImp(psMovNro, False)
If Not rs.EOF And Not rs.BOF Then
    If rs!Moneda = gMonedaExtranjera Then
        lsSimbolo = "$."
    Else
        lsSimbolo = "S/."
    End If
    sTexto = sTexto + BON & PrnSet("N+") & "DESCOMPOSICION DE EFECTIVO : SALIDA" & BOFF & PrnSet("N-") & Chr$(10)
    sTexto = sTexto + IIf(pbCondensado, Con, "") + String(pnColPage - 3, "=") & Chr$(10)
    nMontoM = 0: nMontoB = 0
    nMontoTot = 0
    sTexto = sTexto + BON & ImpreFormat(" BILLETES ", 22, 7) & ImpreFormat("MONTO", 15) _
         & ImpreFormat(" MONEDAS ", 14, 5) & ImpreFormat("MONTO", 10) + BOFF & Chr$(10)
    sTexto = sTexto + String(pnColPage - 3, "-") + BOFF & Chr$(10)
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
        sTexto = sTexto + lsBilletes + lsMonedas & Chr$(10)
    Next
    sTexto = sTexto + String(pnColPage - 3, "-") & Chr$(10)
    sTexto = sTexto + BON & ImpreFormat("SUBTOTAL :", 12, 3) & ImpreFormat(lsSimbolo, 3, 1) & ImpreFormat(nMontoB, 15, 2, True) _
         & ImpreFormat("SUBTOTAL", 10, 8) & ImpreFormat(lsSimbolo, 3, 1) & ImpreFormat(nMontoM, 10, 2, True) + BOFF & Chr$(10)
    sTexto = sTexto + String(pnColPage - 3, "=") & Chr$(10)
    sTexto = sTexto + Space(45) & BON & "TOTAL :  " _
           & lsSimbolo & " " & PrnVal(nMontoTot, 14, 2) & BOFF & Chr$(10)
    sTexto = sTexto + Space(45) & String(31, "=") + COFF & Chr$(10)
    sTexto = sTexto + "" & Chr$(10)
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
         Print #ArcSal, Chr$(18);   'Retorna al tipo de letra normal
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
    & "         JOIN RRHH   RH ON RH.CPERSCOD = P.CPERSCOD " _
    & "         LEFT JOIN AREAS  A ON A.CAREACOD = RH.CAREACOD " _
    & "         LEFT JOIN AGENCIAS AG ON AG.CAGECOD = RH.cAgenciaActual " _
    & " WHERE   P.CPERSCOD ='" & psPersCod & "'"

Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    psAgeCod = rs!cAgeCod
    psAgeDesc = rs!cAgeDescripcion
    psAreaCod = rs!cAreaCod
    psAreaDesc = rs!CAREADESCRIPCION
    psPersNombre = PstaNombre(rs!cPersNombre)
Else
    Exit Function
End If
rs.Close
Set rs = Nothing

BON = PrnSet("B+")
BOFF = PrnSet("B-")
Con = PrnSet("C+")
COFF = PrnSet("C-")

On Error GoTo ImprimeReciboIngresoEgresoErr
    nAncho = pnColPage
    sTexto = ""
    If pbHabilitacionCH = False Then
        If pbIngreso Then
            sTexto = sTexto + ImpreCabAsiento(pnColPage, pdFecha, psEmpresaLogo, psOpeCod, psMovNro, " R E C I B O   D E   I N G R E S O ", , , , 1) & Chr$(10)
        Else
            sTexto = sTexto + ImpreCabAsiento(pnColPage, pdFecha, psEmpresaLogo, psOpeCod, psMovNro, " R E C I B O   D E   E G R E S O ", , , , 1) & Chr$(10)
        End If
    Else
        sTexto = sTexto + ImpreCabAsiento(pnColPage, pdFecha, psEmpresaLogo, psOpeCod, psMovNro, "H A B I L I T A C I O N    D E    C A J A    C H I C A", , , , 1) & Chr$(10)
    End If
    If pbEsCajachica Then
        sTexto = sTexto + "  CAJA CHICA   : " & BON & ImpreFormat(psCHNro, 12) & ImpreFormat(psCHDesc, 40) & BOFF & Chr$(10)
    End If
    If pnArendirFase = ArendirRendicion Or pnArendirFase = ArendirSustentacion Then
        sTexto = sTexto + "  Rec. Arendir : " & BON & ImpreFormat(psDocNroARendir, 14) & BOFF & "Fecha : " & BON & psFechaDoc & BOFF & Chr$(10)
    End If
    If pnTipoArendir = gArendirTipoViaticos Then
       sTexto = sTexto + "  Rec.Viaticos : " & BON & psNroRecViaticos & BOFF & Chr$(10)
    End If
    sTexto = sTexto + "  Agencia      : " & BON & ImpreFormat(psAgeCod, 14) & psAgeDesc & BOFF & Chr$(10)
    sTexto = sTexto + "  Area         : " & BON & ImpreFormat(psAreaCod, 14) & psAreaDesc & BOFF & Chr$(10)
    sTexto = sTexto + "  Persona      : " & BON & ImpreFormat(psPersCod, 14) & psPersNombre & BOFF & Chr$(10)
    sTexto = sTexto + "  Importe      : " & BON & ConvNumLet(lnMonto, True, False) & BOFF & Chr$(10)
    sTexto = sTexto + ImpreGlosa(psGlosa, pnColPage, "  Concepto     : ") & Chr$(10)
    'sTexto = sTexto + "" & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10)
    'sTexto = sTexto + BON & Space(20) & "_________________________" & Chr$(10)
    'Texto = sTexto + Space(20) & Centra("Vo Bo. CAJA ", 25) & BOFF & Chr$(10) & Chr$(10) & Chr$(10) & Chr(10)
    If pbHabilitacionCH = False Then
        sTexto = sTexto + ImprePiePag(pnColPage, "1")
    Else
        sTexto = sTexto + ImprePiePag(pnColPage, "98", "Vo Bo GERENCIA")
    End If
    'End If
    ImprimeReciboIngresoEgreso = sTexto & Chr(12) & sTexto
    Exit Function
    
ImprimeReciboIngresoEgresoErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeReciboIngresoEgreso Method")
End Function
Public Function ImprimePlanillaRendicion(ByVal psMovNroSol As String, ByVal pnColPage As Integer, ByVal pnLinPage As Integer, ByVal psNomCmac As String, _
                                         ByVal pdFecSis As Date, ByVal psCtaContArendir As String, ByVal psCtaContPendiente As String, ByVal psSimbolo As String, _
                                         ByVal pnMoneda As Moneda, ByVal pnTipoArendir As ArendirTipo) As String
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
Dim Con As String
Dim COFF As String
Dim lnSaldo As Currency
Dim lnSust As Currency

BON = PrnSet("B+")
BOFF = PrnSet("B-")
Con = PrnSet("C+")
COFF = PrnSet("C-")


P = 0
nLin = 66
Cont = 0
Set oNARendir = New NARendir
Set rs = New ADODB.Recordset
Set rs = oNARendir.GetDatosArendir(psMovNroSol, psCtaContArendir, psCtaContPendiente, pnMoneda, pnTipoArendir)
If Not rs.EOF And Not rs.BOF Then
    lnSaldo = rs!nMovSaldo
    lnSust = 0
    'Identificamos los Documentos Sustentatorios del Recibo de A rendir
    Do While Not rs.EOF
        lsTexto = lsTexto + CabeceraPlanillaRendicion(pnTipoArendir, " DECLARACION DE A RENDIR CUENTA ", nLin, P, pnLinPage, pnColPage, psNomCmac, pdFecSis, rs!cDocNro, IIf(IsNull(rs!cDocDescAtenc), "ATENCION", rs!cDocDescAtenc), _
                IIf(IsNull(rs!cDocNroAtenc), "EFECTIVO", rs!cDocNroAtenc), IIf(IsNull(rs!dFechaDocAtenc), Mid(rs!CMOVATENC, 7, 2) & "/" & Mid(rs!CMOVATENC, 5, 2) & "/" & Mid(rs!CMOVATENC, 1, 4), rs!dFechaDocAtenc), psSimbolo, rs!ImporteARendir, rs!cAgeCod, rs!cAgeDescripcion, _
                rs!cAreaCod, rs!CAREADESCRIPCION, rs!cPersCod, rs!cPersNombre, rs!cAreaCodArendir, rs!cAgeCodArendir, rs!cAreaArendir, rs!cAgeArendir)

        lsTexto = lsTexto + Con & "   " & Format(rs.Bookmark, "000") & "  " & rs!dDocFechaSust & "  " & _
                 Left(rs!cDocNroAbrevSust & Space(4), 4) & " " & Left(rs!cDocNroSust & Space(14), 14) & " " & _
                 Left(PstaNombre(rs!cPersCodSust) & Space(40), 40) & _
                 "  " & Left(rs!cDescSust & Space(35), 35) & " " & Right(Space(14) & Format(rs!nDocImporte, "#,##0.00"), 14) & COFF & Chr$(10)
                 
        nLin = nLin + 1
        lnSust = lnSust + rs!nDocImporte
        rs.MoveNext
    Loop
    lsTexto = lsTexto + Con
    lsTexto = lsTexto + " ------------------------------------------------------------------------------------------------------------------------------------" & Chr$(10)
    lsTexto = lsTexto + "                                                                              TOTAL SUSTENTADO                 :  " & psSimbolo & " " & PrnVal(CCur(Format(lnSust, "##,###0.00")), 14, 2) & Chr$(10)
    lsTexto = lsTexto + "" & Chr$(10)
    lsTexto = lsTexto + BON + "                                                                              SALDO POR RENDIR EN CAJA         :  " & psSimbolo & " " & PrnVal(CCur(Format(lnSaldo, "##,###0.00")), 14, 2) + BOFF + Chr$(10)
    lsTexto = lsTexto + " ====================================================================================================================================" & Chr$(10)
    lsTexto = lsTexto + COFF + Chr$(10) + Chr$(10) + Chr$(10) + Chr$(10) + Chr$(10) + Chr$(10)
    lsTexto = lsTexto + Centra("______________________    ______________________   ______________________", pnColPage) & Chr$(10)
    lsTexto = lsTexto + Centra("   Vo Bo  USUARIO           Vo Bo JEFE DE AREA         Vo Bo GERENCIA    ", pnColPage) & Chr$(10) & Chr$(10)
End If
rs.Close
Set rs = Nothing
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
Dim Con As String
Dim COFF As String

If nLin > pnLinPage - 4 Then
    BON = PrnSet("B+")
    BOFF = PrnSet("B-")
    Con = PrnSet("C+")
    COFF = PrnSet("C-")
   If P > 0 Then lsTexto = lsTexto & Chr$(12)
   P = P + 1
   lsTexto = lsTexto + ImpreFormat(psNomCmac, 60) & pdFecSis & " - " & Format(Time, "hh:mm:ss") + Chr$(10)
   lsTexto = lsTexto + " CONTABILIDAD    " & Space(55) & "Pag. " & Format(P, "000") + Chr$(10)
   lsTexto = lsTexto + BON & Centra(psTitulo, pnColPage) & BOFF + Chr$(10) + Chr$(10)
   If pnTipoArendir = gArendirTipoCajaGeneral Then
        lsTexto = lsTexto + Con & "  Recibo de A rendir: " & BON & psNroDocArendir & BOFF + Chr$(10)
        lsTexto = lsTexto + "  A Rendir de.......: " & BON & psAreaCodARendir + psAgeCod & " " & psAreaDescArendir + " " + psAgeDescArendir & BOFF + Chr$(10)
        lsTexto = lsTexto + "  " & Mid(psDescDocAtend & Space(18), 1, 18) & ": " & BON & psNroDocAtend & BOFF & "       Fecha : " & BON & psDocAtencFecha & BOFF & _
            Space(15) & "Monto : " & BON & Right(Space(18) & psSimbolo & " " & Format(ImporteARendir, "#,#0.00"), 18) & BOFF + Chr$(10)
   Else
        lsTexto = lsTexto + Con & "  Recibo de A rendir: " & BON & psNroDocArendir & BOFF + Space(15) & "Monto : " & BON & Right(Space(18) & psSimbolo & " " & Format(ImporteARendir, "#,#0.00"), 18) & BOFF + Chr$(10)
        lsTexto = lsTexto + "  A Rendir de.......: " & BON & psAreaCodARendir + psAgeCodArendir & " " & psAreaDescArendir + " " + psAgeDescArendir & BOFF + Chr$(10)
   End If
   lsTexto = lsTexto + "  Usuario...........: " & BON & psPersCod & " " & PstaNombre(psPersNombre) & BOFF + Chr$(10)
   lsTexto = lsTexto + "  Area Funcional....: " & BON & psAreaCod & " " & psAreaDesc & BOFF + Chr$(10)
   If psAgeCod <> "" Then
        lsTexto = lsTexto + "  Agencia       ....: " & BON & psAgeCod & " " & psAgeDesc & BOFF + Chr$(10)
   End If
   lsTexto = lsTexto + Chr$(10)
   If pbCabRend Then
        lsTexto = lsTexto + " ====================================================================================================================================" & Chr$(10)
        lsTexto = lsTexto + "        DOCUMENTOS SUSTENTATORIOS          " & Chr$(10)
        lsTexto = lsTexto + "  ITEM  --------------------------------  PROVEEDOR                                 DETALLE                                  IMPORTE " & Chr$(10)
        lsTexto = lsTexto + "        Fecha      Tpo  Número                                                                                                       " & Chr$(10)
        lsTexto = lsTexto + " ------------------------------------------------------------------------------------------------------------------------------------" & COFF + Chr$(10)
    Else
        lsTexto = lsTexto + COFF + BON & "ASIENTO CONTABLE EN " & IIf(psSimbolo = "S/.", "M.N. :", "M.E. :") & BOFF & Chr$(10)
        lsTexto = lsTexto + String(pnColPage - 1, "=") & Chr$(10)
        lsTexto = lsTexto + "  " & BON & "Cuenta Contable" & Space(43) & "DEBE       HABER" & BOFF & Chr$(10)
        lsTexto = lsTexto + String(pnColPage - 1, "-") & Chr$(10) & Chr$(10)
    End If
    nLin = 15
End If
CabeceraPlanillaRendicion = lsTexto
End Function
Public Function ImprimeAsientoSustArendir(ByVal psMovNroAtenc As String, ByVal pnColPage As Integer, ByVal pnLinPage As Integer, ByVal psNomCmac As String, _
                                         ByVal pdFecSis As Date, ByVal psCtaContArendir As String, ByVal psSimbolo As String, ByVal pnTipoArendir As ArendirTipo)
Dim lsTexto As String
Dim rs As ADODB.Recordset
Dim oArendir As NARendir

Dim oGen As DGeneral
Dim BON As String: Dim BOFF As String
Dim Con As String:   Dim COFF As String

Dim lnDebe As Currency: Dim lnHaber As Currency
Dim nLin As Integer
Dim P As Integer


Set oGen = New DGeneral

Set oArendir = New NARendir
Set rs = New ADODB.Recordset

BON = PrnSet("B+")
BOFF = PrnSet("B-")
Con = PrnSet("C+")
COFF = PrnSet("C-")
nLin = 66
P = 0
Set rs = oArendir.GetAsientosSustARendir(psMovNroAtenc, psCtaContArendir, pnTipoArendir)
If Not rs.EOF And Not rs.BOF Then
    Do While Not rs.EOF
        
        lsTexto = lsTexto + CabeceraPlanillaRendicion(pnTipoArendir, "SUSTENTACION DE A RENDIR CUENTA ", nLin, P, pnLinPage, pnColPage, psNomCmac, pdFecSis, rs!cDocNro, IIf(rs!cDocDescAtenc = "", "ATENCION", rs!cDocDescAtenc), _
               IIf(rs!cDocNroAtenc = "", "EFECTIVO", rs!cDocNroAtenc), IIf(IsNull(rs!dFechaDocAtenc), Mid(rs!CMOVATENC, 7, 2) & "/" & Mid(rs!CMOVATENC, 5, 2) & "/" & Mid(rs!CMOVATENC, 1, 4), rs!dFechaDocAtenc), psSimbolo, rs!ImporteARendir, rs!cAgeCod, rs!cAgeDescripcion, _
               rs!cAreaCod, rs!CAREADESCRIPCION, rs!cPersCod, rs!cPersNombre, rs!cAreaCodArendir, rs!cAgeCodArendir, rs!cAreaArendir, rs!cAgeArendir, False)
               
        lsTexto = lsTexto + Con + "   " & Justifica(rs!CtaContSust, 18) & "  " & Justifica(oGen.CuentaNombre(rs!CtaContSust), 66) & " " & psSimbolo & " " & PrnVal(rs!debe, 14, 2) & " " & psSimbolo & " " & PrnVal(rs!haber, 14, 2) & Chr$(10)
        lnDebe = lnDebe + rs!debe
        lnHaber = lnHaber + rs!haber
        nLin = nLin + 1
        rs.MoveNext
    Loop
    lsTexto = lsTexto + COFF + Chr$(10) + String(pnColPage - 1, "-") & Chr$(10)
    lsTexto = lsTexto + Con
    lsTexto = lsTexto + Space(75) & "TOTALES        " & psSimbolo & Right(Space(15) & Format(lnDebe, "###,###,##0.00"), 15) & "  " & psSimbolo & Right(Space(15) & Format(lnHaber, "###,###,##0.00"), 15) & Chr$(10)
    lsTexto = lsTexto + COFF
    lsTexto = lsTexto + String(pnColPage - 1, "=") & Chr$(10)
    lsTexto = lsTexto + ImprePiePag(pnColPage, "89", "VoBo  SOLICITANTE")
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
Dim Con As String:   Dim COFF As String
Dim lnDebe As Currency: Dim lnHaber As Currency
Dim lsCtaCont As String

Set oGen = New DGeneral
Set oArendir = New NARendir
Set rs = New ADODB.Recordset

BON = PrnSet("B+")
BOFF = PrnSet("B-")
Con = PrnSet("C+")
COFF = PrnSet("C-")

lsTexto = ""

Set rs = oArendir.GetDetDocCtaCont(psMovNroSol, psCtaContArendir, psCtaContPendiente, pnTipoArendir)

If Not rs.EOF And Not rs.BOF Then
    sCabe = Con + String(125, "=") & Chr(10) _
      & "  D O C U M E N T O               C O N C E P T O                                                DEBE           HABER" & Chr(10) _
      & "TIPO  NUMERO        FECHA " & Chr(10) _
      & String(125, "-") + COFF & Chr(10)

    nLi = pnLinPage: P = 0
    lsTexto = PrnSet("MI", 4)
    If nLi > pnLinPage - 6 Then
        lsTexto = lsTexto + Cabecera("DETALLE DE DOCUMENTOS POR CUENTA CONTABLE", P, psSimbolo, pnColPage, sCabe) & Chr$(10)
        nLi = 9
    End If
    Do While Not rs.EOF
        lsTexto = lsTexto + Con + "CUENTA CONTABLE: " & Justifica(rs!CtaContSust, 18) & " " & Justifica(oGen.CuentaNombre(rs!CtaContSust), 80) & Chr$(10) & Chr$(10)
        nLi = nLi + 2
        lsCtaCont = rs!CtaContSust
        lnDebe = 0: lnHaber = 0
        Do While lsCtaCont = rs!CtaContSust
            lnDebe = lnDebe + rs!debe
            lnHaber = lnHaber + rs!haber
            lsTexto = lsTexto + Justifica(rs!cDocAbrev, 3) & " " & Justifica(rs!cDocNro, 16) & " " & Justifica(rs!dDocFecha, 10) & "  " & Justifica(rs!cMovDesc, 60) & " " & PrnVal(rs!debe, 14, 2) & " " & PrnVal(rs!haber, 14, 2) & Chr(10)
            nLi = nLi + 1
            If nLi > pnLinPage - 6 Then
                lsTexto = lsTexto + Cabecera("DETALLE DE DOCUMENTOS POR CUENTA CONTABLE", P, psSimbolo, pnColPage, sCabe) & Chr$(10)
                nLi = 9
            End If
            rs.MoveNext
            If rs.EOF Then Exit Do
        Loop
        lsTexto = lsTexto + String(125, "-") & Chr$(10)
        lsTexto = lsTexto + Space(94) & PrnVal(lnDebe, 14, 2) & " " & PrnVal(lnHaber, 14, 2) & Chr$(10) & Chr$(10)
        nLi = nLi + 3
    Loop
End If
rs.Close
Set rs = Nothing
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
Dim Con As String
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
Con = PrnSet("C+")
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
lsTexto = Con + BON + PrnSet("I+") + ImpreFormat(lsTitulo, lnCol, 50) & ImpreFormat(lsTitulo, lnCol, 10) + PrnSet("I-") + BOFF + Chr$(10) + Chr$(10)
lsTexto = lsTexto + ImpreFormat(Trim(psAgeDesc), lnCol, 3) & ImpreFormat(Trim(psAgeDesc), lnCol, lnColIzq) + Chr$(10)
lnLineas = lnLineas + 1
lsTexto = lsTexto + ImpreFormat(CentrarCadena(String(33, "=") + " Concepto " + String(35, "="), lnCol), lnCol, 3) + ImpreFormat(CentrarCadena(String(33, "=") + " Concepto " + String(35, "="), lnCol), lnCol, lnColIzq) & Chr(10)
lnLineas = lnLineas + 1
'Impresion de la glosa
lsGlosa = JustificaTexto(psGlosa, lnCol - 6)
lsGlosaLin = "": lnPos = 0: Char = ""
lnLineasGlosa = 0
For I = 1 To Len(lsGlosa)
    Char = Mid(lsGlosa, I, 1)
    lsGlosaLin = lsGlosaLin & Char
    If Char = Chr(10) Or I = Len(lsGlosa) Then
        lsGlosaLin = Mid(lsGlosaLin, 1, IIf(I = Len(lsGlosa), Len(lsGlosaLin), Len(lsGlosaLin) - 1))
        lsTextoGlosa = lsTextoGlosa & ImpreFormat(Trim(lsGlosaLin), lnCol, lnColIzq - 7) & ImpreFormat(Trim(lsGlosaLin), lnCol, lnColIzq) & Chr(10)
        lnLineas = lnLineas + 1
        lnLineasGlosa = lnLineasGlosa + 1
        If lnLineasGlosa = 3 Then
            Exit For
        End If
        lsGlosaLin = ""
    End If
Next
lsTexto = lsTexto + lsTextoGlosa
lsTexto = lsTexto + ImpreFormat(String(50, "_") & " CARGOS " & String(10, "_") & " ABONOS " & String(14, "_"), lnCol, 3) + ImpreFormat(String(50, "_") & " CARGOS " & String(10, "_") & " ABONOS " & String(14, "_"), lnCol, lnColIzq) & Chr$(10)
If lnImporteCargo > 0 Then
    lsTexto = lsTexto + ImpreFormat(ImpreFormat("IMPORTE :", 40) & ImpreFormat(lnImporteCargo, 14, 2, True) & ImpreFormat(0, 15, 2, True), lnCol, 3) + ImpreFormat(ImpreFormat("IMPORTE :", 40) & ImpreFormat(lnImporteCargo, 14, 2, True) & ImpreFormat(0, 15, 2, True), lnCol, lnColIzq) & Chr$(10)
Else
    lsTexto = lsTexto + ImpreFormat(ImpreFormat("IMPORTE :", 40) & ImpreFormat(0, 14, 2, True) & ImpreFormat(lnImporteAbono, 15, 2, True), lnCol, 3) + ImpreFormat(ImpreFormat("IMPORTE :", 40) & ImpreFormat(0, 14, 2, True) & ImpreFormat(lnImporteAbono, 15, 2, True), lnCol, lnColIzq) & Chr$(10)
End If
lsTexto = lsTexto + ImpreFormat(String(lnCol, "_"), lnCol, 3) + ImpreFormat(String(lnCol, "_"), lnCol, lnColIzq) & Chr$(10)
lsTexto = lsTexto + ImpreFormat("Sr(es) :", lnCol, 3) + ImpreFormat("Sr(es) :", lnCol, lnColIzq) & Chr$(10)
lsTexto = lsTexto + BON + ImpreFormat(PstaNombre(psPersNombre), lnCol, lnColIzq + 5) + ImpreFormat(PstaNombre(psPersNombre), lnCol, lnColIzq + 5) & Chr$(10)
lsTexto = lsTexto + ImpreFormat(psPersDireccion, lnCol, lnColIzq + 5) + ImpreFormat(psPersDireccion, lnCol, lnColIzq + 5) & Chr$(10)
lsTexto = lsTexto + ImpreFormat(psPersUbiGeo, lnCol, lnColIzq + 5) + ImpreFormat(psPersUbiGeo, lnCol, lnColIzq + 5) + BOFF + Chr$(10) & Chr(10)

lsTexto = lsTexto + ImpreFormat(String(3, "-") & "FECHA" & String(13, "-") & " MONEDA " & String(5, "-") & " CUENTA " & String(15, "-") + BON + lsTextoTotal + BOFF & String(15, "-"), lnCol + 4, 3) _
            + ImpreFormat(String(3, "-") & "FECHA" & String(13, "-") & " MONEDA " & String(5, "-") & " CUENTA " & String(15, "-") + BON + lsTextoTotal + BOFF & String(8, "-"), lnCol + 4, lnColIzq) & Chr$(10)

lsTexto = lsTexto + ImpreFormat(String(3, " ") & Format(pdFecha & " " & Time, "mm/dd/yyyy hh:mm:ss") & String(3, " ") & IIf(psMoneda = gMonedaNacional, "SOLES", "DOLARES") & String(3, " ") & psNroCtaAhorros + BON + ImpreFormat(pnImporte, 15, 2, True) + BOFF + ImpreFormat(psCodUser, 5, 4), lnCol + 4, 3) _
            + ImpreFormat(String(3, " ") & Format(pdFecha & " " & Time, "mm/dd/yyyy hh:mm:ss") & String(3, " ") & IIf(psMoneda = gMonedaNacional, "SOLES", "DOLARES") & String(3, " ") & psNroCtaAhorros + BON + ImpreFormat(pnImporte, 15, 2, True) + BOFF + ImpreFormat(psCodUser, 5, 4), lnCol + 4, lnColIzq) & Chr$(10)

lsTexto = lsTexto + ImpreFormat(String(lnCol, "-"), lnCol, 3) + ImpreFormat(String(lnCol, "-"), lnCol, lnColIzq) + COFF + Chr$(10)

'ImprimeNotaCargoAbono = lsTexto & Chr(12) & lsTexto
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
   Linea sTexto, "  " & prs!cMovnro & "  " & Format(prs!nMovItem, "000") & "  " & prs!cCtaContCod & "   " & PrnVal(prs!nDebe, 15, 2) & PrnVal(prs!nHaber, 15, 2)
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
            sTexto = sTexto + Chr$(12)
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
        sTexto = sTexto + PrnSet("MI", 5) & ImpreCabAsiento(pnColPage, pdFecSis, psEmpresaLogo, psOpeCod, lsMovNro, lsTitulo) & Chr$(10)
        sTexto = sTexto + "  Documento Refe.  : " & BON & Mid(psNroDocRef, 1, 50) & BOFF & Space(10) & "Fecha : " & BON & pdFechaDocRef & BOFF & Chr$(10)
        sTexto = sTexto + "  Area             : " & BON & psAreaCod & "  " & ImpreCarEsp(psAreaDesc) & BOFF & Chr$(10)
        sTexto = sTexto + "  Agencia          : " & BON & psAgeCod & "  " & ImpreCarEsp(psAgeDesc) & BOFF & Chr$(10)
        sTexto = sTexto + "  Persona          : " & BON & psPersCod & "  " & ImpreCarEsp(psPersNombre) & BOFF & Chr$(10)
        sTexto = sTexto + "  Cargo            : " & BON & Mid(psPersCargo, 1, 50) & BOFF & Chr$(10)
        sTexto = sTexto + "  Categoria        : " & BON & Mid(psCategoria, 1, 50) & BOFF & Chr(10)
        sTexto = sTexto + "  Destino          : " & BON & Justifica(Mid(!Destino, 1, 50), 50) & BOFF & Chr(10)
        sTexto = sTexto + "  Transporte Ida   : " & BON & Justifica(Mid(!Ida, 1, 50), 50) & BOFF & Chr(10)
        sTexto = sTexto + "  Transporte Vuelta: " & BON & Justifica(Mid(!Vuelta, 1, 50), 50) & BOFF & Chr$(10) & Chr$(10)
        sTexto = sTexto + "  Partida      : " & BON & !Partida & BOFF & Space(10) & "Nro. Días    : " & BON & PrnVal(!Dias, 5, 0) & BOFF & Chr(10)
        sTexto = sTexto + "  Importe      : " & BON & ConvNumLet(PrnVal(!Importe, 10, 2), True, False) & BOFF & Chr(10)
        lsGlosa = !Motivo
        sTexto = sTexto + ImpreGlosa(lsGlosa, pnColPage, "  Concepto     : ") & Chr(10)
        sTexto = sTexto + "   ==================================================================" & Chr(10)
        sTexto = sTexto + "     C O N C E P T O                                   IMPORTE (" & psSimbolo & ")" & Chr(10)
        sTexto = sTexto + "   ------------------------------------------------------------------" & Chr$(10)
        lnTotal = 0
        Do While Not rsConc.EOF
            lnImporte = 0
            For I = 1 To rsAux.Fields.Count - 1
                If rsAux.Fields(I).Name = Trim(rsConc!concepto) Then
                    lnImporte = rsAux.Fields(I).value
                    Exit For
                End If
            Next
            sTexto = sTexto + "      " & Justifica(rsConc!Descripcion, 45) & "  " & PrnVal(lnImporte, 14, 2) & Chr(10)
            lnTotal = lnTotal + lnImporte
            rsConc.MoveNext
        Loop
        sTexto = sTexto + "   ------------------------------------------------------------------" & Chr$(10)
        sTexto = sTexto + BON + "                                   T O T A L     " & psSimbolo & " " & PrnVal(lnTotal, 14, 2) + BOFF + Chr$(10)
        sTexto = sTexto + "   ==================================================================  " & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10)
        sTexto = sTexto + Space(20) & Centra("_____________________________", 50) & Chr$(10)
        sTexto = sTexto + Space(20) & BON & Centra(Mid(psPersNombre, 1, 49), 50) & Chr$(10)
        sTexto = sTexto + Space(20) & Centra("  L.E. " & psDocPers, 50) & BOFF & Chr$(10) & Chr$(10)
        
        lsGlosa = "De acuerdo a la Política existente, la rendición del presente recibo se deberá realizar en el plazo establecido, bajo responsabilidad." & Chr$(10) _
               & "Los Comprobantes de gastos serán a nombre de " & BON & Trim(psEmpresa) & BOFF & ", RUC " & BON & Trim(psEmpresaRuc) & BOFF & Chr$(10) & " "
        sTexto = sTexto + ImpreGlosa(lsGlosa, pnColPage, " NOTA : ") & Chr$(10) & Chr$(10) & Chr$(10)
        sTexto = sTexto + Centra("____________________          _____________________", 80) & Chr$(10)
        sTexto = sTexto + Centra("    Vo Bo RRHH                    Vo Bo GERENCIA    ", 80) & Chr$(10)
        sTexto = sTexto + Chr$(10) & Chr$(10) & Chr$(10) & PrnSet("MI", 0)
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
Dim nAncho As Integer, N As Integer
Dim lsGlosa As String
Dim I As Integer
Dim lnImporte As Currency
Dim lsMovNro As String
Dim lnTotal As Currency
Dim rs As ADODB.Recordset
Dim Rsc As ADODB.Recordset
Dim oArendir As NARendir
Dim lsTitulo As String

Set oArendir = New NARendir
Set rs = New ADODB.Recordset
Set Rsc = New ADODB.Recordset

sTexto = ""
nAncho = pnColPage
Set rs = oArendir.GetDatosTotalesViaticos(psMovNroViat)
sTexto = ""
With rs
    Do While Not .EOF
        lsMovNro = rs!cMovnro
        If sTexto <> "" Then
            sTexto = sTexto + Chr$(12)
        End If
        lsTitulo = "RECIBO DE VIATICOS A RENDIR CUENTA NRO. " & !DOCVIAT
        sTexto = sTexto + PrnSet("MI", 5) & ImpreCabAsiento(pnColPage, pdFecSis, psEmpresaLogo, !copecod, lsMovNro, lsTitulo) & Chr$(10)
        sTexto = sTexto + "  Documento Refe.  : " & BON & Mid(!NRODOCREF, 1, 50) & BOFF & Space(10) & "Fecha : " & BON & !dDocFecha & BOFF & Chr$(10)
        sTexto = sTexto + "  A Rendir de...   : " & BON & !cCodAreaArendir & !cCodAgeArendir & "  " & ImpreCarEsp(!cAreaArendir) & " " & ImpreCarEsp(!cAgeArendir) & BOFF & Chr$(10)
        sTexto = sTexto + "  Persona          : " & BON & !cPersCod & "  " & ImpreCarEsp(PstaNombre(!cPersNombre)) & BOFF & Chr$(10)
        sTexto = sTexto + "  Area             : " & BON & !cAreaCod & "  " & ImpreCarEsp(!CAREADESCRIPCION) & BOFF & Chr$(10)
        sTexto = sTexto + "  Agencia          : " & BON & !cAgeCod & "  " & ImpreCarEsp(!cAgeDescripcion) & BOFF & Chr$(10)
        sTexto = sTexto + "  Cargo            : " & BON & Mid(!crhcargodescripcion, 1, 50) & BOFF & Chr$(10)
        sTexto = sTexto + "  Categoria        : " & BON & Mid(!Categoria, 1, 50) & BOFF & Chr(10)
        sTexto = sTexto + "  Destino          : " & BON & Justifica(Mid(!Destino, 1, 50), 50) & BOFF & Chr(10)
        sTexto = sTexto + "  Transporte Ida   : " & BON & Justifica(Mid(!TransIda, 1, 50), 50) & BOFF & Chr(10)
        sTexto = sTexto + "  Transporte Vuelta: " & BON & Justifica(Mid(!TrasnVuelta, 1, 50), 50) & BOFF & Chr$(10) & Chr$(10)
        sTexto = sTexto + "  Partida      : " & BON & !dPartida & BOFF & Space(10) & "Nro. Días    : " & BON & PrnVal(!nMovViaticosDias, 5, 0) & BOFF & Chr(10)
        sTexto = sTexto + "  Importe      : " & BON & ConvNumLet(rs!copecod, PrnVal(!nMovMonto, 10, 2), False) & BOFF & Chr(10)
        lsGlosa = !Motivo
        sTexto = sTexto + ImpreGlosa(lsGlosa, pnColPage, "  Concepto     : ") & Chr(10)
        sTexto = sTexto + "   ==================================================================" & Chr(10)
        sTexto = sTexto + "     C O N C E P T O                                   IMPORTE (" & psSimbolo & ")" & Chr(10)
        sTexto = sTexto + "   ------------------------------------------------------------------" & Chr$(10)
        lnTotal = 0
        Set Rsc = oArendir.GetDetalleCostosViaticos(rs!cMovnro)
        Do While Not Rsc.EOF
            sTexto = sTexto + "      " & Justifica(Rsc!Descripcion, 45) & "  " & PrnVal(Rsc!Importe, 14, 2) & Chr(10)
            lnTotal = lnTotal + Rsc!Importe
            Rsc.MoveNext
        Loop
        Rsc.Close
        Set Rsc = Nothing
        sTexto = sTexto + "   ------------------------------------------------------------------" & Chr$(10)
        sTexto = sTexto + BON + "                                   T O T A L     " & psSimbolo & " " & PrnVal(lnTotal, 14, 2) + BOFF + Chr$(10)
        sTexto = sTexto + "   ==================================================================  " & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10)
        sTexto = sTexto + Space(20) & Centra("_____________________________", 50) & Chr$(10)
        sTexto = sTexto + Space(20) & BON & Centra(Mid(PstaNombre(!cPersNombre), 1, 49), 50) & Chr$(10)
        sTexto = sTexto + Space(20) & Centra("  LE/DNI " & !DNI, 50) & BOFF & Chr$(10) & Chr$(10)
        
        lsGlosa = "De acuerdo a la Política existente, la rendición del presente recibo se deberá realizar en el plazo establecido, bajo responsabilidad." & Chr$(10) _
               & "Los Comprobantes de gastos serán a nombre de " & BON & Trim(psEmpresa) & BOFF & ", RUC " & BON & Trim(psEmpresaRuc) & BOFF & Chr$(10) & " "
        sTexto = sTexto + ImpreGlosa(lsGlosa, pnColPage, " NOTA : ") & Chr$(10) & Chr$(10) & Chr$(10)
        sTexto = sTexto + Centra("____________________          _____________________", 80) & Chr$(10)
        sTexto = sTexto + Centra("    Vo Bo RRHH                    Vo Bo GERENCIA    ", 80) & Chr$(10)
        sTexto = sTexto + Chr$(10) & Chr$(10) & Chr$(10) & PrnSet("MI", 0)
        If .EOF Then
            Exit Do
        End If
        .MoveNext
    Loop
    ImprimirReciboViaticoData = sTexto
End With
   
End Function
Public Sub inicio(psEmpNom As String, psAgeNom As String, psFecSis As String)
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


Set rs = New ADODB.Recordset
Set oCh = New NCajaChica


P = 0
nLin = 66
Cont = 0
BON = PrnSet("B+")
BOFF = PrnSet("B-")
Con = PrnSet("C+")
COFF = PrnSet("C-")

sTexto = PrnSet("MI", 5) & CabeceraRepo(psSimbolo, pnLinPage, psNomCmact, pdFecSis, pnColPage, _
                                         pbValido, psNroCH, pnNroProc, psCHDesc, _
                                         pdFechaApertReembolso, pnMontoAsignado, psPersCod, psPersNombre, nLin, P)
                                         
Linea sTexto, Con & BON & " EGRESOS CON A RENDIR : " & BOFF
Linea sTexto, " --------------------   " & COFF
nLin = nLin + 2
nTotalR = 0:   nTotalI = 0:   nTotalE = 0: nSaldoR = 0: lnTotalSust = 0
If Not rsArendir Is Nothing Then
    Do While Not rsArendir.EOF
        sTexto = sTexto & CabeceraRepo(psSimbolo, pnLinPage, psNomCmact, pdFecSis, pnColPage, _
                                       pbValido, psNroCH, pnNroProc, psCHDesc, _
                                       pdFechaApertReembolso, pnMontoAsignado, psPersCod, psPersNombre, nLin, P)
        nLin = nLin + 2
        Linea sTexto, Con & "  Recibo de A rendir: " & rsArendir![Nro Doc] & "   " & rsArendir!Fecha & "   " & _
              Space(15) & "Monto : " & Right(Space(18) & psSimbolo & " " & Format(rsArendir!Importe, "#,#0.00"), 18) & ImpreFormat(IIf(Val(rsArendir!NroCajaCh) = pnNroProc, "", " Caja Chica N° " & rsArendir!NroCajaCh), 20, 10) & COFF
        Linea sTexto, Con & "  Usuario...........: " & rsArendir!Solicitante & COFF
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
              sTexto = sTexto & CabeceraRepo(psSimbolo, pnLinPage, psNomCmact, pdFecSis, pnColPage, _
                                             pbValido, psNroCH, pnNroProc, psCHDesc, _
                                             pdFechaApertReembolso, pnMontoAsignado, psPersCod, psPersNombre, nLin, P)
                                             
              Linea sTexto, Con + PrnSet("I+") & ImpreFormat(Format(Cont, "000"), 3, 5) & ImpreFormat(rs!FECHADOC, 10) & _
                       ImpreFormat(rs!cDocAbrev, 3) & ImpreFormat(rs!cDocNro, 15) & ImpreFormat(PstaNombre(rs!cPersNombre), 35) & _
                       ImpreFormat(rs!Glosa, 32) & ImpreFormat(rs!Importe, 12, 2, True) & PrnSet("I-") + COFF
              lnTotalSust = lnTotalSust + Abs(rs!Importe)
              rs.MoveNext
           Loop
        End If
        rs.Close
        Set rs = Nothing
        If Val(rsArendir!Saldo) = 0 Then
            Set rs = oCh.GetDatosRendArendirCH(rsArendir!cMovNroAtenc, psCtaFondoFijo)
            If Not rs.EOF Then
               Cont = Cont + 1
               nLin = nLin + 1
               CabeceraRepo psSimbolo, pnLinPage, psNomCmact, pdFecSis, pnColPage, _
                             pbValido, psNroCH, pnNroProc, psCHDesc, _
                             pdFechaApertReembolso, pnMontoAsignado, psPersCod, psPersNombre, nLin, P
                             
               Linea sTexto, Con + BON + ImpreFormat(Format(Cont, "000"), 3, 5) & ImpreFormat(rs!FechaRend, 10) & _
                            ImpreFormat(IIf(rs!nMovImporte > 0, "R/I ", "R/E"), 3) & ImpreFormat(rs!cMovnro, 52) & ImpreFormat(IIf(rs!nMovImporte > 0, "VUELTO INGRESADO A CAJA", "REEMBOLSO A USUARIO"), 30) & " " & _
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
           Linea sTexto, Con + PrnSet("I+") + ImpreFormat(Format(Cont, "000"), 74, 5) & _
                ImpreFormat(IIf(nImporte < 0, "REEMBOLSO PENDIENTE A USUARIO", "RENDICION PENDIENTE"), 30) & " " & _
                ImpreFormat(nImporte, 13, 2, True) + PrnSet("I-") + COFF
           nLin = nLin + 1
        End If
        rsArendir.MoveNext
    Loop

    Linea sTexto, Con & " ------------------------------------------------------------------------------------------------------------------------------------" & COFF
    Linea sTexto, Con & "   " & Space(60) & "TOTAL EGRESOS SUSTENTADOS    " & psSimbolo & " " & PrnVal(lnTotalSust, 16, 2)    ' nTotalR - nTotalI + nTotalE
    Linea sTexto, Con & "   " & Space(60) & "TOTAL INGRESOS POR RENDICION " & psSimbolo & " " & PrnVal(nTotalI, 16, 2)
    Linea sTexto, Con & "   " & Space(60) & "TOTAL EGRESOS  POR RENDICION " & psSimbolo & " " & PrnVal(nTotalE, 16, 2)
    Linea sTexto, Con & "   " & Space(60) & "---------------------------------------------------------------------"
    Linea sTexto, Con & "   " & Space(80) & "TOTAL EGRESOS POR A RENDIR   " & psSimbolo & " " & PrnVal(nTotalR, 16, 2)
    Linea sTexto, Con & "   " & Space(80) & "TOTAL A RENDIR PENDIENTES    " & psSimbolo & " " & PrnVal(nSaldoR, 16, 2)
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
          Linea sTexto, Con & "   " & Format(Cont, "000") & " " & Justifica(rsEgreDir!Tipo, 4) & " " & Justifica(rsEgreDir!Fecha, 10) & " " & _
                             Justifica(rsEgreDir![Nro Doc], 16) & " " & Justifica(PstaNombre(rsEgreDir!Solicitante), 35) & " " & Justifica(rsEgreDir!concepto, 41) & " " & PrnVal(rsEgreDir!Importe, 14, 2) & COFF
          nTotal = nTotal + rsEgreDir!Importe
       rsEgreDir.MoveNext
    Loop
    Linea sTexto, Con & " ------------------------------------------------------------------------------------------------------------------------------------" & COFF
    Linea sTexto, Con & "   " & Space(80) & "TOTAL EGRESOS DIRECTOS       " & psSimbolo & " " & PrnVal(nTotal, 16, 2), 2
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
   If P > 0 Then sCabe = sCabe & Chr$(12)
   P = P + 1
   Linea sCabe, COFF + ImpreFormat(psNomCmact, 60, 0) & Space(42) & pdFecSis & " - " & Format(Time, "hh:mm:ss")
   Linea sCabe, "     CONTABILIDAD    " & Space(55) & "Pag. " & Format(P, "000")
   Linea sCabe, BON & Centra(" DOCUMENTOS SUSTENTATORIOS DE CAJA CHICA ", pnColPage) & BOFF
   Linea sCabe, Con
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
Dim rs As ADODB.Recordset
Dim oCh As NCajaChica
Dim oGen As DGeneral
Dim lsImp As String
Dim lnTotalDebe As Currency
Dim lnTotalHaber As Currency
Dim nLi    As Integer, P As Integer
Dim sCabe  As String

Set oGen = New DGeneral
Set rs = New ADODB.Recordset
Set oCh = New NCajaChica

Linea sAsiento, ImpreCabAsiento(pnColPage, pdFecSis, psEmpresaLogo, psOpeCod, psMovNro, "", True, Format(pdFecSis, "dd/mm/yyyy"), False, 2)
If pbValido = False Then
   Linea sAsiento, BON & Centra("*** REPORTE NO VALIDO PARA RENDICION EN CONTABILIDAD ***", 80) & BOFF, 2
End If

Set rs = oCh.GetAsientoCH(psAreaCh, psAgeCh, pnProcCH)
lnTotalDebe = 0
lnTotalHaber = 0
sTexto = ""
Do While Not rs.EOF
   Linea sTexto, ImpreFormat(rs!cCtaContCod, 19, 5) & "  " & ImpreFormat(oGen.CuentaNombre(rs!cCtaContCod), 60) & ImpreFormat(psSimbolo, 3) & ImpreFormat(rs!debe, 12, 2, True) & ImpreFormat(psSimbolo, 3) & ImpreFormat(rs!haber, 12, 2, True)
   lnTotalDebe = lnTotalDebe + rs!debe
   lnTotalHaber = lnTotalHaber + rs!haber
   rs.MoveNext
Loop
rs.Close
Set rs = Nothing
lsImp = sAsiento & Chr(10) & ImpreAsiento(pnColPage, psSimbolo, sTexto, lnTotalDebe, lnTotalHaber, True, "ENCARGADO")

''Ahora Imprimimos Detalle de Documentos por Cuenta en el Debe
Set rs = oCh.GetDocumentosAsientoCH(psAreaCh, psAgeCh, pnProcCH, psCtaFondoFijo)
sCabe = Con + String(125, "=") & Chr(10) _
      & "  D O C U M E N T O               C O N C E P T O                                                DEBE           HABER" & Chr(10) _
      & "TIPO  NUMERO        FECHA " & Chr(10) _
      & String(125, "-") & Chr(10)

nLi = pnLinPage: P = 0
sTexto = PrnSet("MI", 4)
Do While Not rs.EOF
   If nLi > pnLinPage - 6 Then
      Linea sTexto, COFF + Cabecera("DETALLE DE GASTOS DE CAJA CHICA POR DOCUMENTO", P, psSimbolo, pnColPage, sCabe)
      nLi = 9
   End If
   Linea sTexto, Con + BON + "CUENTA CONTABLE: " & ImpreFormat(rs!cCtaContCod, 18) & ImpreFormat(oGen.CuentaNombre(rs!cCtaContCod), 80) + BOFF, 2
   nLi = nLi + 2
   sAsiento = rs!cCtaContCod
   lnTotalDebe = 0: lnTotalHaber = 0
   Do While sAsiento = rs!cCtaContCod
      lnTotalDebe = lnTotalDebe + rs!debe
      lnTotalHaber = lnTotalHaber + rs!haber
      Linea sTexto, ImpreFormat(rs!cDocAbrev, 3, 0) & ImpreFormat(rs!cDocNro, 16) & ImpreFormat(rs!dDocFecha, 10) & ImpreFormat(rs!cMovDesc, 50) & ImpreFormat(rs!debe, 14, 2, True) & ImpreFormat(rs!haber, 14, 2, True)
      nLi = nLi + 1
      If nLi > pnLinPage - 6 Then
         Linea sTexto, COFF + Cabecera("DETALLE DE GASTOS DE CAJA CHICA POR DOCUMENTO", P, psSimbolo, pnColPage, sCabe)
         nLi = 9
      End If
      rs.MoveNext
      If rs.EOF Then Exit Do
   Loop
   Linea sTexto, String(125, "-")
   Linea sTexto, BON + ImpreFormat("", 40) & ImpreFormat("TOTAL", 40) & ImpreFormat(lnTotalDebe, 15, 2, True) & " " & ImpreFormat(lnTotalHaber, 13, 2, True) + BOFF, 2
   nLi = nLi + 3
Loop
lsImp = lsImp & Chr(12) & Con & sTexto & COFF

ImprimeAsientoRendicionCH = lsImp


End Function
Public Function ImpreAsiento(ByVal pnColPage As Integer, ByVal psSimbolo As String, _
                            ByVal psAsiento As String, ByVal pnTotalDebe As Currency, _
                            ByVal pnTotalHaber As Currency, lPiePage As Boolean, Optional lsArea As String = "CAJA") As String
Dim sImpre As String, nLon  As String
sImpre = sImpre & Chr$(10)
' ASIENTO CONTABLE
sImpre = sImpre & BON & "ASIENTO CONTABLE EN " & IIf(psSimbolo = "S/.", "M.N. :", "M.E. :") & BOFF & Chr$(10)
sImpre = sImpre & String(pnColPage - 1, "=") & Chr$(10)
sImpre = sImpre & "  " & BON & "Cuenta Contable" & Space(43) & "DEBE       HABER" & BOFF & Chr$(10)
sImpre = sImpre & String(pnColPage - 1, "-") & Chr$(10)
sImpre = sImpre & Con & psAsiento & COFF     'Detalle de Asiento
sImpre = sImpre & String(pnColPage - 1, "-") & Chr$(10)
sImpre = sImpre & Con
sImpre = sImpre & Space(75) & "TOTALES        " & psSimbolo & Right(Space(15) & Format(pnTotalDebe, "###,###,##0.00"), 15) & "   " & psSimbolo & Right(Space(15) & Format(pnTotalHaber, "###,###,##0.00"), 15) & Chr$(10)
sImpre = sImpre & COFF
sImpre = sImpre & String(pnColPage - 1, "=") & Chr$(10)

If lPiePage Then
   sImpre = sImpre & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10)
   sImpre = sImpre & BON
   sImpre = sImpre & Space(5) & "_____________________                           ______________________" & Chr$(10)
   lsArea = "Vo Bo " & lsArea
   nLon = Int((21 - Len(lsArea)) / 2)
   sImpre = sImpre & Space(5) & Space(nLon) & lsArea & Space(nLon) & "                              Vo Bo CONTABILIDAD" & BOFF & Chr$(10)
End If
ImpreAsiento = sImpre
End Function

Public Function ImprimeDocSalidaEfectivo(ByVal pnColPage As Integer, ByVal psFecha As String, _
                ByVal psOpeCod As String, ByVal psMovNro As String, ByVal psNomCmact As String, Optional ByVal lnNumCop As Long = 1) As String
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
lsImpre = lsImpre & ImpreCabAsiento(pnColPage, CDate(psFecha), psNomCmact, psOpeCod, psMovNro, psTit1, True) & Chr$(10)
Linea lsImpre, ImpreFormat("Area/Agencia Origen ", 25) & ":" & ImpreFormat(lsAreaOrigen, 40)
Linea lsImpre, ImpreFormat("Area/Agencia Destino", 25) & ":" & ImpreFormat(lsAreaDestino, 40), 2
Linea lsImpre, BON + ImpreFormat("TRANSPORTE", 25) + BOFF
If lsNombreTransp <> "" Then
    Linea lsImpre, ImpreFormat("TIPO ", 20, 5) & ":" & ImpreFormat(lsTipoTransporte, 40)
    Linea lsImpre, ImpreFormat("EMPRESA", 20, 5) & ":" & ImpreFormat(lsNombreTransp, 50)
    Linea lsImpre, ImpreFormat("DOCUMENTO", 20, 5) & ":" & ImpreFormat(IIf(lsNroDocTransp = "", "NO INDICADO!!!", lsNroDocTransp), 20)
    Linea lsImpre, BON + ImpreFormat("COSTO TRANSPORTE", 20, 5) & ":  S/." & ImpreFormat(lnImporteTrans, 10, 2, True) + BOFF, 3
Else
    Linea lsImpre, ImpreFormat("NO INDICADO", 20, 5), 3
End If
lsBilletaje = ImprimeBilletaje(Mid(psOpeCod, 3, 1), psMovNro, pnColPage, False)
Linea lsImpre, lsBilletaje
'Linea lsImpre, ImprePiePag(pnColPage, "158", "Responsable Traslado")
Linea lsImpre, ImprePiePag(pnColPage, IIf(lsNroDocSal = "", "158", "1"), IIf(lsNroDocSal = "", "Responsable Traslado", ""))

Dim lsAux As String
For I = 1 To lnNumCop - 1
    lsAux = lsAux & Chr$(12) & lsImpre
Next
lsImpre = lsImpre & lsAux
ImprimeDocSalidaEfectivo = lsImpre

End Function

Public Function ImprimeDocConfHabEfectivo(ByVal pnColPage As Integer, ByVal psFecha As String, _
                                        ByVal psOpeCod As String, ByVal psMovNro As String, ByVal psNomCmact As String)
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
lsImpre = lsImpre & ImpreCabAsiento(pnColPage, CDate(psFecha), psNomCmact, psOpeCod, psMovNro, psTit1, True) & Chr$(10)
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
            ByVal psNomAge As String, ByVal psMovNro As String, ByVal psLpt As String, Optional ByVal pscodcmac As String)

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

lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)
lnEspaceIni = 2
lnLineasBol = 22
lnAnchoBol = 45

nFicSal = FreeFile
Open psLpt For Output As nFicSal
Print #nFicSal, Chr$(27) & Chr$(64);
Print #nFicSal, Chr$(27) & Chr$(108) & Chr$(0)
Print #nFicSal, Chr$(27) + Chr$(50);   'espaciamiento lineas 1/6 pulg.
Print #nFicSal, Chr$(27) + Chr$(67) + Chr$(lnLineasBol);  'Longitud de página a 22 líneas'
Print #nFicSal, Chr$(27) + Chr$(77);   'Tamaño 10 cpi
Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita


lsFecha = Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4)
'20010703 1347191120100EJRS
lsHora = Format$(Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2), "hh:mm:ss")
lsUserOrig = Right(psMovNro, 4)

If pscodcmac = "102" Then
    Print #nFicSal, ImpreCarEsp(PrnSet("B+") + ImpreFormat("CMCPL - OPERACIONES", lnAnchoBol, lnEspaceIni))
    Print #nFicSal, ImpreCarEsp(ImpreFormat(Trim(psNomAge) + IIf(pnMoneda = gMonedaNacional, "-SOLES", "-DOLARES"), lnAnchoBol, lnEspaceIni) + PrnSet("B-"))
    Print #nFicSal, ImpreCarEsp(ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni))
    Print #nFicSal, ""
    Print #nFicSal, ImpreCarEsp(ImpreFormat("PARA :[" & psUserDest & "] " & psNomUserDest, lnAnchoBol, lnEspaceIni))
    Print #nFicSal, ImpreCarEsp(ImpreFormat("DE   :[" & psUserOrig & "] " & psNomUserOrigen, lnAnchoBol, lnEspaceIni))
    Print #nFicSal, ""
    'Print #nFicSal, ImpreCarEsp(PrnSet("B+") + CentrarCadena(psTitCab, lnAnchoBol - 10) & CentrarCadena(psTitCab, lnAnchoBol))
    Print #nFicSal, ImpreCarEsp(PrnSet("B+") + Space(lnEspaceIni) + CentrarCadena(psTitCab, lnAnchoBol - 3))
    Print #nFicSal, ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni) + PrnSet("B-"))
    Print #nFicSal, ImpreCarEsp(ImpreFormat(ImpreFormat(psTitDet, lnAnchoBol / 2) & ImpreFormat(CCur(pnImporte), 15, 2, True), lnAnchoBol, lnEspaceIni))
    For I = 1 To 3
        Print #nFicSal, ""
    Next
    Print #nFicSal, ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni))
    Print #nFicSal, ImpreCarEsp(ImpreFormat(lsUserOrig, lnAnchoBol, lnEspaceIni))
Else
    Print #nFicSal, ImpreCarEsp(PrnSet("B+") + ImpreFormat("CMACT - OPERACIONES", lnAnchoBol, lnEspaceIni) & ImpreFormat("CMACT - OPERACIONES", lnAnchoBol, lnEspaceIni))
    Print #nFicSal, ImpreCarEsp(ImpreFormat(Trim(psNomAge) + IIf(pnMoneda = gMonedaNacional, "-SOLES", "-DOLARES"), lnAnchoBol, lnEspaceIni) & ImpreFormat(Trim(psNomAge) + IIf(pnMoneda = gMonedaNacional, "-SOLES", "-DOLARES"), lnAnchoBol, lnEspaceIni) + PrnSet("B-"))
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
End If
Dim oGen As DGeneral
Set oGen = New DGeneral
lsMensaje = oGen.GetMensajeBoletas("")
Set oGen = Nothing
If pscodcmac = "102" Then
    Print #nFicSal, PrnSet("B+") + ImpreCarEsp(ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) + PrnSet("B-"))
Else
Print #nFicSal, PrnSet("B+") + ImpreCarEsp(ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) & ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) + PrnSet("B-"))
End If
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
   nOPE = 0
   Do While Not Rope.EOF
      DoEvents
      nOPE = nOPE + 1
      mMat(nOPE, 0) = Format(Rope!nOpeNiv, "000")
      mMat(nOPE, 1) = Rope!copecod
      mMat(nOPE, 2) = Rope!cOpeDesc
      sOpe = Rope!copecod
      
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
       Linea sTexto, String(215, "-") + Chr(12)
       nLin = 1
    End If
Next
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
           Linea sTexto, String(lnCarLin, "-") & Chr$(12)
        End If
        lnPagNro = lnPagNro + 1
        ' Definición de Cabecera 1
        lsCabe01 = vsEmpNom
        lsCabe01 = lsCabe01 & Space((lnCarLin - 34) - Len(lsCabe01))
        lsCabe01 = lsCabe01 & "PAGINA: " & FillNum(Trim(Str(lnPagNro)), 4, " ")
        lsCabe01 = lsCabe01 & Space(5) & "FECHA: " & Format(vsFecSis, "dd/mm/yyyy")
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
prs.Close: Set prs = Nothing
Linea lsImpre, String(lnCarLin, "-")
ImprimeObjetos = lsImpre
End Function

Public Sub ImprimeBoletaCompraVenta(ByVal psTitCab As String, ByVal psTitDet As String, _
                ByVal psNomPers As String, ByVal psDirPers As String, _
                ByVal psDocumento As String, ByVal pnTipoCambio As Double, _
                ByVal psCodOpe As String, ByVal pnMontoDol As Currency, ByVal pnMontoSoles As Currency, _
                ByVal psNomAge As String, ByVal psMovNro As String, ByVal psLpt As String, Optional ByVal pscodcmac As String)
Dim sCad As String
Dim nFicSal As Integer
Dim lsFecha As String
Dim lsHora As String
Dim lsUserOrig As String
Dim sBoldON As String
Dim sBoldOFF As String
Dim lsMensaje As String * 39
Dim I As Integer
Dim lnNumLineas As Integer
Dim lnLineasBol As Integer
Dim lnEspaceIni As Integer
Dim lnAnchoBol As Long
Dim psSubTit As String
Dim psSubTit2 As String
Dim sAgencia As String * 25, sNumMov As String * 25
Dim lnMovNro As Long
Dim oMov As DMov
Set oMov = New DMov
lnMovNro = oMov.GetnMovNro(psMovNro)
Set oMov = Nothing

ETIQ:
On Error GoTo ERROR

sBoldON = Chr$(27) + Chr$(71)
sBoldOFF = Chr$(27) + Chr$(72)


sCad = sCad & Chr$(27) & Chr$(64)
sCad = sCad & Chr$(27) + Chr$(50)   'espaciamiento lineas 1/6 pulg.
sCad = sCad & Chr$(27) + Chr$(67) + Chr$(22)  'Longitud de página a 22 líneas'
sCad = sCad & Chr$(27) + Chr$(77)   'Tamaño 10 cpi
sCad = sCad & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
sCad = sCad & Chr$(27) + Chr$(18) ' cancela condensada
sCad = sCad & Chr$(27) + Chr$(72) ' desactiva negrita

Select Case psCodOpe
    Case gOpeCajeroMECompra, gOpeCajeroMECompraEsp
        psTitDet = "Monto Comprado $"
        psSubTit = "Neto a Pagar S/."
        psSubTit2 = "Tipo de Cambio Compra"
    Case gOpeCajeroMEVenta, gOpeCajeroMEVentaEsp
        psTitDet = "Monto Vendido $"
        psSubTit = "Neto a Recibir S/."
        psSubTit2 = "Tipo de Cambio Venta"
End Select

lsFecha = Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4)
lsHora = Format$(Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2), "hh:mm:ss")
lsUserOrig = Right(psMovNro, 4)
lnEspaceIni = 2
lnLineasBol = 22
lnAnchoBol = 45

If pscodcmac = "102" Then
    sCad = sCad & Chr$(10)
    sCad = sCad & ImpreCarEsp(sBoldON & ImpreFormat("CM - OPERACIONES", lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(ImpreFormat(psNomAge, 30) & ImpreFormat(lnMovNro, 10, 0, False), lnAnchoBol, lnEspaceIni) & sBoldOFF) & Chr$(10) '& ImpreFormat(ImpreFormat(psNomAge, 30) & ImpreFormat(lnMovNro, 10, 0, False), lnAnchoBol, lnEspaceIni) & sBoldOFF) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni)) & Chr$(10) ' & ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(psNomPers, lnAnchoBol, lnEspaceIni)) & Chr$(10) ' & ImpreFormat(psNomPers, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat("Dir:" & psDirPers, lnAnchoBol, lnEspaceIni)) & Chr$(10) ' & ImpreFormat("Dir:" & psDirPers, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    If psDocumento <> "" Then
        sCad = sCad & ImpreCarEsp(ImpreFormat("Nro Doc:" & psDocumento, lnAnchoBol, lnEspaceIni)) & Chr$(10) ' & ImpreFormat("Nro Doc:" + psDocumento, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    Else
        sCad = sCad & Chr$(10)
    End If
    sCad = sCad & ImpreCarEsp(sBoldON & Space(lnEspaceIni) + CentrarCadena(psTitCab, lnAnchoBol - 3)) & Chr$(10) ' & CentrarCadena(psTitCab, lnAnchoBol - 3, lnEspaceIni + 4)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni)) & Chr$(10) ' & ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(ImpreFormat(psTitDet, lnAnchoBol / 2) & ImpreFormat(CCur(pnMontoDol), 15, 4, True), lnAnchoBol, lnEspaceIni) & sBoldOFF) & Chr$(10) ' & ImpreFormat(ImpreFormat(psTitDet, lnAnchoBol / 2) & ImpreFormat(CCur(pnMontoDol), 15, 2, True), lnAnchoBol, lnEspaceIni) & sBoldOFF) & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(ImpreFormat(psSubTit, lnAnchoBol / 2) & ImpreFormat(CCur(pnMontoSoles), 15, 4, True), lnAnchoBol, lnEspaceIni)) & Chr$(10) ' & ImpreFormat(ImpreFormat(psSubTit, lnAnchoBol / 2) & ImpreFormat(CCur(pnMontoSoles), 15, 2, True), lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(ImpreFormat(psSubTit2, lnAnchoBol / 2) & ImpreFormat(CCur(pnTipoCambio), 15, 4, True), lnAnchoBol, lnEspaceIni)) & Chr$(10) '& ImpreFormat(ImpreFormat(psSubTit2, lnAnchoBol / 2) & ImpreFormat(CCur(pnTipoCambio), 15, 4, True), lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni)) & Chr$(10) '& ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(Right(psMovNro, 4), lnAnchoBol, lnEspaceIni)) & Chr$(10) '' & ImpreFormat(Right(psMovNro, 4), lnAnchoBol, lnEspaceIni)) & Chr$(10)
Else
    sCad = sCad & Chr$(10)
    sCad = sCad & ImpreCarEsp(sBoldON & ImpreFormat("CMACI - OPERACIONES", lnAnchoBol, lnEspaceIni) & ImpreFormat("CMACI - OPERACIONES", lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(ImpreFormat(psNomAge, 30) & ImpreFormat(lnMovNro, 10, 0, False), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat(psNomAge, 30) & ImpreFormat(lnMovNro, 10, 0, False), lnAnchoBol, lnEspaceIni) & sBoldOFF) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(psNomPers, lnAnchoBol, lnEspaceIni) & ImpreFormat(psNomPers, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat("Dir:" & psDirPers, lnAnchoBol, lnEspaceIni) & ImpreFormat("Dir:" & psDirPers, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    If psDocumento <> "" Then
        sCad = sCad & ImpreCarEsp(ImpreFormat("Nro Doc:" & psDocumento, lnAnchoBol, lnEspaceIni) & ImpreFormat("Nro Doc:" + psDocumento, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    Else
        sCad = sCad & Chr$(10)
    End If
    sCad = sCad & ImpreCarEsp(sBoldON & Space(lnEspaceIni) + CentrarCadena(psTitCab, lnAnchoBol - 3) & CentrarCadena(psTitCab, lnAnchoBol - 3, lnEspaceIni + 4)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni) & ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(ImpreFormat(psTitDet, lnAnchoBol / 2) & ImpreFormat(CCur(pnMontoDol), 15, 4, True), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat(psTitDet, lnAnchoBol / 2) & ImpreFormat(CCur(pnMontoDol), 15, 4, True), lnAnchoBol, lnEspaceIni) & sBoldOFF) & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(ImpreFormat(psSubTit, lnAnchoBol / 2) & ImpreFormat(CCur(pnMontoSoles), 15, 4, True), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat(psSubTit, lnAnchoBol / 2) & ImpreFormat(CCur(pnMontoSoles), 15, 4, True), lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(ImpreFormat(psSubTit2, lnAnchoBol / 2) & ImpreFormat(CCur(pnTipoCambio), 15, 4, True), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat(psSubTit2, lnAnchoBol / 2) & ImpreFormat(CCur(pnTipoCambio), 15, 4, True), lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni) & ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(Right(psMovNro, 4), lnAnchoBol, lnEspaceIni) & ImpreFormat(Right(psMovNro, 4), lnAnchoBol, lnEspaceIni)) & Chr$(10)

End If
Dim oGen As DGeneral
Set oGen = New DGeneral
lsMensaje = oGen.GetMensajeBoletas("")
Set oGen = Nothing
sCad = sCad & sBoldON & ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) & ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) & sBoldOFF & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & Chr$(10)

nFicSal = FreeFile
Open psLpt For Output As nFicSal
Print #nFicSal, sCad
Close nFicSal
Exit Sub
ERROR:
    Close nFicSal
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub ImprimeBoletaExtornos(ByVal psTitCab As String, ByVal psTitDet As String, _
        ByVal psCodOpe As String, ByVal pnImporte As Double, ByVal psUserOpe As String, _
        ByVal psNomPers As String, ByVal psNomAge As String, ByVal psMovNro As String, _
        ByVal psNomCmac As String, ByVal nMoneda As Moneda, ByVal psLpt As String, _
        Optional lnMovNro As Long = 0, Optional ByVal pscodcmac As String)

Dim nFicSal As Integer, sCad As String
Dim lsFecha As String
Dim lsHora As String
Dim lsUserOrig As String
Dim sBoldON As String
Dim sBoldOFF As String
Dim lsMensaje As String * 39
Dim I As Integer
Dim lnNumLineas As Integer
Dim lnLineasBol As Integer
Dim lnEspaceIni As Integer
Dim lnAnchoBol As Long
Dim psSubTit As String
Dim psSubTit2 As String
Dim lsMoneda As String

ETIQ:
On Error GoTo ERROR

sBoldON = Chr$(27) + Chr$(71)
sBoldOFF = Chr$(27) + Chr$(72)


sCad = sCad & Chr$(27) & Chr$(64)
sCad = sCad & Chr$(27) & Chr$(50)  'espaciamiento lineas 1/6 pulg.
sCad = sCad & Chr$(27) & Chr$(67) & Chr$(22) 'Longitud de página a 22 líneas'
sCad = sCad & Chr$(27) & Chr$(77)  'Tamaño 10 cpi
sCad = sCad & Chr$(27) + Chr$(107) + Chr$(0)    'Tipo de Letra Sans Serif
sCad = sCad & Chr$(27) + Chr$(18) ' cancela condensada
sCad = sCad & Chr$(27) + Chr$(72) ' desactiva negrita

lsFecha = Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4)

lsHora = Format$(Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2), "hh:mm:ss")
lsUserOrig = Right(psMovNro, 4)
lnEspaceIni = 2
lnLineasBol = 22
lnAnchoBol = 50

If nMoneda <> gMonedaNacional And nMoneda <> gMonedaExtranjera Then
    lsMoneda = ""
Else
    lsMoneda = IIf(nMoneda = gMonedaNacional, "-SOLES", "-DOLARES")
End If

If pscodcmac = "102" Then
    sCad = sCad & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & ImpreCarEsp(sBoldON & ImpreFormat(psNomCmac & " - OPERACIONES", lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(ImpreFormat(psNomAge, 30) & ImpreFormat(lsMoneda, 10) & ImpreFormat(lnMovNro, 10, 0, False), lnAnchoBol, lnEspaceIni) & sBoldOFF) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(psNomPers, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat("Operación de :  " & psUserOpe, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & ImpreCarEsp(sBoldON & Space(lnEspaceIni) + CentrarCadena(psTitCab, lnAnchoBol - 3)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(ImpreFormat("MONTO EXTORNADO ", lnAnchoBol / 2) & ImpreFormat(CCur(pnImporte), 15, 2, True), lnAnchoBol, lnEspaceIni) & sBoldOFF) & Chr$(10)
    
   
    sCad = sCad & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & Chr$(27) & Chr$(15) + ImpreFormat(psTitDet, lnAnchoBol + 25, lnEspaceIni + 2) & Chr$(27) & Chr$(15) & Chr$(27) & Chr$(77) & Chr$(10)
    sCad = sCad & Chr$(27) & Chr$(64)
    sCad = sCad & Chr$(27) & Chr$(77)  'Tamaño 10 cpi
    sCad = sCad & Chr$(27) + Chr$(107) + Chr$(0)    'Tipo de Letra Sans Serif
    sCad = sCad & ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(Right(psMovNro, 4), lnAnchoBol, lnEspaceIni)) & Chr$(10)

Else
    sCad = sCad & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & ImpreCarEsp(sBoldON & ImpreFormat(psNomCmac & " - OPERACIONES", lnAnchoBol, lnEspaceIni) & ImpreFormat(psNomCmac & " - OPERACIONES", lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(ImpreFormat(psNomAge, 30) & ImpreFormat(lsMoneda, 10) & ImpreFormat(lnMovNro, 10, 0, False), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat(psNomAge, 20) & ImpreFormat(lsMoneda, 10) & ImpreFormat(lnMovNro, 10, 0, False), lnAnchoBol, lnEspaceIni) & sBoldOFF) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(psNomPers, lnAnchoBol, lnEspaceIni) & ImpreFormat(psNomPers, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat("Operación de :  " & psUserOpe, lnAnchoBol, lnEspaceIni) & ImpreFormat("Operación de :  " & psUserOpe, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & ImpreCarEsp(sBoldON & Space(lnEspaceIni) + CentrarCadena(psTitCab, lnAnchoBol - 3) & CentrarCadena(psTitCab, lnAnchoBol - 3, lnEspaceIni + 4)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni) & ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(ImpreFormat("MONTO EXTORNADO ", lnAnchoBol / 2) & ImpreFormat(CCur(pnImporte), 15, 2, True), lnAnchoBol, lnEspaceIni) & ImpreFormat(ImpreFormat("MONTO EXTORNADO", lnAnchoBol / 2) & ImpreFormat(CCur(pnImporte), 15, 2, True), lnAnchoBol, lnEspaceIni) & sBoldOFF) & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & ImpreCarEsp(PrnSet("C+") + ImpreFormat(psTitDet, lnAnchoBol + 25, lnEspaceIni + 2) & ImpreFormat(psTitDet, lnAnchoBol + 25, lnEspaceIni + 10) + PrnSet("C-")) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni) & ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & ImpreCarEsp(ImpreFormat(Right(psMovNro, 4), lnAnchoBol, lnEspaceIni) & ImpreFormat(Right(psMovNro, 4), lnAnchoBol, lnEspaceIni)) & Chr$(10)

End If
Dim oGen As DGeneral
Set oGen = New DGeneral
lsMensaje = oGen.GetMensajeBoletas("")
Set oGen = Nothing
If pscodcmac = "102" Then
    sCad = sCad & sBoldON & ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) & sBoldOFF & Chr$(10)
Else
    sCad = sCad & sBoldON & ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) & ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) & sBoldOFF & Chr$(10)
End If
sCad = sCad & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & Chr$(10)

nFicSal = FreeFile
Open psLpt For Output As nFicSal
Print #nFicSal, sCad
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

lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)

nFicSal = FreeFile
Open psLpt For Output As nFicSal

Print #nFicSal, Chr$(27) & Chr$(64);

Print #nFicSal, Chr$(27) + Chr$(50);   'espaciamiento lineas 1/6 pulg.
Print #nFicSal, Chr$(27) + Chr$(67) + Chr$(lnLineasBol);  'Longitud de página a 22 líneas'
Print #nFicSal, Chr$(27) + Chr$(77);   'Tamaño 10 cpi
Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita

lsFecha = Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4)
'20010703 1347191120100EJRS
lsHora = Format$(Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2), "hh:mm:ss")
lsUserOrig = Right(psMovNro, 4)
lnEspaceIni = 2
lnLineasBol = 22
lnAnchoBol = 50

If Mid(psCodOpe, 3, 1) <> gMonedaNacional And Mid(psCodOpe, 3, 1) <> gMonedaExtranjera Then
    lsMoneda = "-SOLES"
Else
    lsMoneda = IIf(Mid(psCodOpe, 3, 1) = gMonedaNacional, "-SOLES", "-DOLARES")
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
                             ByVal psMovNro As String, ByVal psOpeCod As String, _
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
lsCab1 = Con + ImpreFormat("Resp. Caja   : [" & psCodPersResp & "] " + psResponsable, (pnNumCol / 2) + 65) + BON + ImpreFormat("Inicio  :" & psHoraIni, 25) + COFF + BOFF
lsCab2 = Con + ImpreFormat("Resp. Arqueo : [" & psCodPers & "] " + psPersNombre, (pnNumCol / 2) + 65) + BON + ImpreFormat("Término :" & psHoraFin, 25) + COFF + BOFF
lsTitulo2 = PrnSet("I+") + "Caja Chica N° [" & psAreaAgeCh & "] " & psCajaChicaNomb + PrnSet("I-")
lnNumPag = 0
Linea lsTexto, CabeRepo(psNomCmact, psAgencia, "CONTABILIDAD", IIf(Mid(psOpeCod, 3, 1) = "1", "SOLES", "DOLARES"), Format(pdFecha, "dd/mm/yyyy"), "A R Q U E O   D E   C A J A   C H I C A", lsTitulo2, "", "", lnNumPag, pnNumCol)
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
                If rsTotales.Bookmark = 5 Or rsTotales.Bookmark = 7 Then
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
            Optional pdFecSis As Date, Optional psNomAge As String = "", Optional psCodUser As String) As String

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

ETIQ:

On Error GoTo ERROR

lnTope = 0 '6 'Tope de lineas en Boleta

lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)

lsNroExt = Str(pnNumExt)

nFicSal = FreeFile
Open sLpt For Output As nFicSal

Print #nFicSal, Chr$(27) & Chr$(64);

Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita

sSep = 15
sIni = 1
sMax = 33
saux = 5


sFecha = Format$(pdFecSis, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")
sMonto = Format$(pnMonto, "#,##0.00")
sSDisp = Format$(pnSaldo, "#,##0.00")
lsSaldoC = Format$(pnSaldoC, "#,##0.00")

lsNomAge = psNomAge

'Print #nFicSal, Chr$(10);
Print #nFicSal, lsNegritaOn; 'Activa Negrita
Print #nFicSal, Tab(sIni); "CMACT - AHORRO"; Space(19 + sSep + saux); "CMACT - AHORRO"

If Mid(psCodCta, 9, 1) = 1 Then
    Print #nFicSal, Tab(sIni); Trim(psNomAge) & "-SOLES"; Space(saux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; Space(sSep); Trim(psNomAge) & "-SOLES"; Space(saux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt;
Else
    Print #nFicSal, Tab(sIni); Trim(psNomAge) & "-DOLARES"; Space(saux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt; Space(sSep); Trim(psNomAge) & "-DOLARES"; Space(saux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt;
End If

If psNomAgeRem = "" Then
    Print #nFicSal, ""
Else
    Print #nFicSal, Tab(sIni); "Ag.Rem: " & Trim(psNomAgeRem); Space(saux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))); "Ag.Rem: " & Trim(psNomAgeRem)
End If

If psComCmac = "XXX" Then
    If psLin3 = "XXX" Then
        Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
    Else
        Print #nFicSal, Tab(sIni); psLin3 & Space(saux + sSep + sMax - Len(psLin3)) & psLin3 & lsNegritaOff;  'Desactiva Negrita
        lnNumLinCmac = 1
    End If
    lnNumLinCmac = 0
Else
    Print #nFicSal, Tab(sIni); "NroDocCmac:" & psComCmac & Space(saux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & "NroDocCmac:" & psComCmac & lsNegritaOff;  'Desactiva Negrita
    lnNumLinCmac = 1
End If

Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(saux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora

'psCliente = PstaNombre(psCliente)

lnCliAux = InStr(1, psCliente, "*", vbTextCompare)

If lnCliAux = 0 Then
    If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(psCliente); Space(saux + sMax + sSep - Len(psCliente)); ImpreCarEsp(psCliente)
Else
    lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
    lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
    
    If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
    If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
    
    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux1); Space(saux + sMax + sSep - Len(lsCliAux1)); ImpreCarEsp(lsCliAux1)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux2); Space(saux + sMax + sSep - Len(lsCliAux2)); ImpreCarEsp(lsCliAux2)
    
    lnCliAux = 1
End If

If pbSaldoInt Or pbCuenta Then
    Print #nFicSal, Tab(sIni); "Cuenta:" & psCodCta; Space(8 + sSep + saux); "Cuenta:" & psCodCta
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
    Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(psTexto)); Space(saux + sSep + sMax - Len(Trim(psTexto))); ImpreCarEsp(Trim(psTexto))
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
    Print #nFicSal, Tab(sIni); ImpreCarEsp(psCodUser); Space(29 + sSep + saux); ImpreCarEsp(psCodUser)
Else
    Print #nFicSal, Tab(sIni); "Loc/Rem"; Space(sMax + saux - Len("Loc/Rem") - 1 - 8); ImpreCarEsp(psCodUser) & "/"; ImpreCarEsp(psCodUsuRem); Space(sSep); "Loc/Rem"; Space(sMax + saux - Len("Loc/Rem") - 1 - 8); ImpreCarEsp(psCodUser) & "/"; ImpreCarEsp(psCodUsuRem)
End If
Dim clsGen As DGeneral
Set clsGen = New DGeneral
lsMensaje = clsGen.GetMensajeBoletas(psCodCta)
Set clsGen = Nothing
Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsMensaje); Space(-1 + sSep); ImpreCarEsp(lsMensaje); lsNegritaOff

lnNumLinCmac = lnNumLinCmac + 1
 
'Print #nFicSal, Chr$(12)
For saux = 1 To (lnTope - lnNumLinCmac)
    Print #nFicSal, ""
Next saux
Close #nFicSal
Exit Function
ERROR:
    Close #nFicSal
    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
        GoTo ETIQ
    End If
End Function


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
         DoEvents
         RaiseEvent BarraProgress(R.Bookmark, "LIBRO DIARIO", "", "Procesando... ", vbBlue)
         sFec = Mid(R!cMovnro, 7, 2) & "/" & Mid(R!cMovnro, 5, 2) & "/" & Mid(R!cMovnro, 1, 4)
         sAge = Mid(R!cMovnro, 18, 2)         'Agencia
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
         If R!nMovImporte > 0 Then
            Linea sCadRep, Space(5) + Justifica(R!cCtaContCod, 20) & " " & Justifica(R!cCtaContDesc, 60) & PrnVal(R!nMovImporte, 16, 2)
            nTotalDocD = nTotalDocD + R!nMovImporte
            nTotalD = nTotalD + R!nMovImporte
         Else
            Linea sCadRep, Space(5) + Justifica(R!cCtaContCod, 20) & " " & Justifica(R!cCtaContDesc, 60) & Space(18) & PrnVal(R!nMovImporte * -1, 16, 2)
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
   Linea sCadCab, COFF & BON & Cabecera("L I B R O   D I A R I O ", P, "", 80, "", Format(dFecFin, "dd/mm/yyyy")), 2
   Linea sCadCab, Con & Space(5) + "CODIGO" + Space(15) + "NOMBRE" + Space(64) + "DEBE" + Space(13) + "HABER", 2
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
   nSaldo = oSdo.GetCtaSaldo(sCtaCod, Format(dFechaDel - 1, "mm/dd/yyyy"))
   Linea sTexto, ImprimeLibroMayorCab(nLin, P, dFechaDel, dFechaAl, sCtaCod, sCtaDes, pnLinPage), 0
   nDebe = 0: nHaber = 0
   nDebeD = 0: nHaberD = 0
   If rsCta.EOF Then
      Exit Do
   End If
   sFecha = Mid(rsCta!cMovnro, 1, 8)

   sTipoCta = RTrim(rsCta!cCtaCaracter)

   Do While rsCta!cCtaContCod = sCtaCod
      DoEvents
      RaiseEvent BarraProgress(rsCta.Bookmark, "LIBRO MAYOR", "", "Procesando... ", vbBlue)
      If P Mod 20 = 0 Then
         sImpre = sImpre & sTexto
         sTexto = ""
      End If
      If sFecha <> Mid(rsCta!cMovnro, 1, 8) Then
         Linea sTexto, BON & Space(15) & "Total Diario  : " & PrnVal(nDebeD, 16, 2, False) & "  " & PrnVal(nHaberD, 16, 2, False) & "  " & PrnVal(nSaldo, 16, 2) & BOFF, 2
         nLin = nLin + 2
         nDebeD = 0: nHaberD = 0
         sFecha = Mid(rsCta!cMovnro, 1, 8)
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
      Linea sTexto, Mid(rsCta!cMovnro, 7, 2) & "/" & Mid(rsCta!cMovnro, 5, 2) & "/" & Mid(rsCta!cMovnro, 1, 4) & " " & PrnVal(rsCta!nMovCorrela, 5, 0) & "   " & Mid(rsCta!cMovnro, 18, 2) & "-" _
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
   Linea sCabe, BON & Cabecera(" L I B R O    M A Y O R ", P, "", 80, , Format(dFechaAl, "dd/mm/yyyy")), 0
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
nSaldo = oSdo.GetCtaSaldo(psCtaCod, Format(pdFechaDel - 1, "mm/dd/yyyy"))
nSaldoIni = nSaldo
Set oSdo = Nothing

Set rsCta = oCont.GetMayorCuenta(psCtaCod, Format(pdFechaDel, "yyyymmdd"), Format(pdFechaAl, "yyyymmdd"), pnImporte, psFiltro)
Set oCont = Nothing
If rsCta.EOF Then
   Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
   Linea sTexto, BON & "CUENTA CONTABLE : " & Justifica(psCtaCod, 16) & " " & psCtaDes, 2
   Linea sTexto, Con & Space(87) & "SALDO INICIAL : " & PrnVal(nSaldo, 16, 2)
   Linea sTexto, Space(87) & "SALDO FINAL   : " & PrnVal(nSaldo, 16, 2) & COFF & BOFF, 1
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
Linea sTexto, BON & "CUENTA CONTABLE : " & Justifica(psCtaCod, 16) & " " & psCtaDes & BOFF, 2
nLin = nLin + 2
nDebe = 0: nHaber = 0
N = 1

sFecha = Mid(rsCta!cMovnro, 1, 8)
Do While Not rsCta.EOF
   DoEvents
   RaiseEvent BarraProgress(rsCta.Bookmark, "Mayor de Cuenta Contable : " & psCtaCod, "", "Generando Reporte... ", vbBlue)
   sMov = rsCta!cMovnro
   If sFecha <> Mid(sMov, 1, 8) Then
      Linea sTexto, Con & BON & Space(87) & "TOTAL DIARIO  : " & PrnVal(nDebeD, 16, 2, False) & "  " & PrnVal(nHaberD, 16, 2, False) & "  " & PrnVal(nSaldo, 16, 2) & COFF & BOFF, 2
      nDebeD = 0: nHaberD = 0
      sFecha = Mid(rsCta!cMovnro, 1, 8)
   End If
   If sTipoCta = "D" Then
      nSaldo = nSaldo + rsCta!nDebe - rsCta!nHaber
   Else
      nSaldo = nSaldo + rsCta!nHaber - rsCta!nDebe
   End If
   
   Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
   'Linea Detalle de Reporte
   sDoc = Justifica(rsCta!cDocAbrev, 3) & " " & Justifica(rsCta!cDocNro, 20) & " " & rsCta!dDocFecha & " "
   sMov = rsCta!cMovnro
   Linea sTexto, Con & " " & Mid(sMov, 1, 8) & "-" & Mid(sMov, 9, 6) & " " & Right(RTrim(sMov), 4) & "  " _
          & " " & sDoc & Mid(Replace(Replace(rsCta!cMovDesc, Chr(13), " "), Chr(10), "") & Space(100), 1, 42) & " " _
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
   Do While sMov = rsCta!cMovnro
      sDoc = Justifica(rsCta!cDocAbrev, 3) & " " & Justifica(rsCta!cDocNro, 20) & " " & rsCta!dDocFecha & " "
      Linea sTexto, Con & " " & Space(23) & sDoc & "  " & COFF
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
Linea sTexto, Con & BON & Space(87) & "TOTAL DIARIO  : " & PrnVal(nDebeD, 16, 2, False) & "  " & PrnVal(nHaberD, 16, 2, False) & "  " & PrnVal(nSaldo, 16, 2) & COFF & BOFF, 2
nLin = nLin + 2
N = N + 1
Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
Linea sTexto, Con & " -------------------------" & String(36, "-") & "------------------------------------------------------------------------------------------------"
Linea sTexto, BON & Space(87) & "TOTAL CUENTA  : " & PrnVal(nDebe, 16, 2, False) & "  " & PrnVal(nHaber, 16, 2, False) & BOFF, 2
nLin = nLin + 2

Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
Linea sTexto, Space(87) & "SALDO INICIAL : " & PrnVal(nSaldoIni, 16, 2)
Linea sTexto, Space(87) & "SALDO FINAL   : " & PrnVal(nSaldo, 16, 2)
Linea sTexto, COFF & BOFF
nLin = pnLinPage + 1
Linea sTexto, Con & " =========================" & String(36, "=") & "================================================================================================" & COFF
ImprimeMayorCta = sTexto
RaiseEvent BarraClose
rs.Close: Set rs = Nothing
rsCta.Close: Set rsCta = Nothing

End Function

Private Function ImprimeMayorCuentaCab(ByRef nLin As Integer, ByRef P As Integer, dFechaDel As Date, dFechaAl As Date, sCtaCod As String, sCtaDes As String, pnLinPage As Integer) As String
Dim k As Integer, nObjs As Integer
Dim sCabe As String
Dim oConec As New DConecta

If nLin > pnLinPage - 6 Then
   oConec.AbreConexion
   Linea sCabe, BON & Cabecera(" M A Y O R   D E   C U E N T A S    C O N T A B L E S", P, "", 80, , Format(oConec.GetFechaHoraServer(), "dd/mm/yyyy")), 0
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
   Do While Not rs.EOF
      If nLin > pnLinPage - 4 Then
         Linea sImpre, CabeRepo(vsEmpNom, vsAgeNom, "Contabilidad", "", vsFecSis, "Configuración de Reportes de Cuentas Contables por Columnas", _
               String(65, "-"), "", "", nPageNro, 79), 0
         nLin = 6
      End If
      If lsOperacion <> Trim(rs!copecod) Then
          Linea sImpre, ""
          Linea sImpre, BON & Trim(rs!copecod) & " - " & Trim(rs!cOpeDesc) & " " & IIf(Mid(rs!copecod, 3, 1) = "1", "MN", "ME") & " (Anexo : " & Trim(Right(rs!copecod, 1)) & ")" & BOFF
          nLin = nLin + 2
      End If
      If lnColumna <> rs!nNroCol Then
          Linea sImpre, Space(5) & rs!nNroCol & "  " & Trim(rs!cDescCol)
          nLin = nLin + 1
      End If
      Linea sImpre, Space(15) & rs!cCtaContCod & " " & Trim(rs!cCtaContDesc)
      nLin = nLin + 1
      lnColumna = Trim(rs!nNroCol)
      lsOperacion = Trim(rs!copecod)
      rs.MoveNext
    Loop
RSClose rs
Set clsRepCtaCol = Nothing
ImprimeRepCtaCol = sImpre
Exit Function
ImprimeRepCtaColErr:
    Call RaiseError(MyUnhandledError, "NContImprimir:ImprimeRepCtaCol Method")
End Function
Public Function ImprimeFlujoCtasBancos(ByVal pnColPage As Integer, ByVal pnLinPage As Integer, ByVal psOpeCod As String, _
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

lsSimbolo = IIf(Mid(psOpeCod, 3, 1) = gMonedaNacional, "S/.", "$.")
Set rs = oCont.GetFlujosCtaIF(psOpeCod, psCtaIfCodDesde, psCtaIfCodHasta, pdDesde, pdHasta)
If rs.EOF And rs.BOF Then
    rs.Close: Set rs = Nothing
    ImprimeFlujoCtasBancos = ""
    Exit Function
End If
RaiseEvent BarraShow(rs.RecordCount)
lsCtaIFCod = ""
lnTotD = 0
lnTotH = 0
Do While Not rs.EOF
    If nLin >= 56 Then
        nLin = pnLinPage
    End If
    CabeceraFlujo lsTexto, nLin, P, pnLinPage, pnColPage, lsSimbolo, pdDesde, pdHasta
    If rs!cPersCod + rs!cCtaIfCod <> lsPersCod + lsCtaIFCod Then
        If lnTotD > 0 Or lnTotH > 0 Then
            Linea lsTexto, Con & " ------------------------------------------------------------------------------------------------------------------------------------"
            nLin = nLin + 1
            Linea lsTexto, BON & Space(75) & "TOTALES  " & ImpreFormat(lnTotD, 14, 2, True) & ImpreFormat(lnTotH, 14, 2, True) & BOFF & COFF, 2
            nLin = nLin + 2
        End If
        lnTotD = 0
        lnTotH = 0
        Linea lsTexto, BON & Con & " CUENTA : [" & rs!cCtaIfCod & "] " & rs!cNomCtaIF & COFF & BOFF
        nLin = nLin + 1
        lnTotD = 0: lnTotH = 0: lnSaldo = 0
        Linea lsTexto, Con & " SALDO ANTERIOR..." & Space(100) & " " & Right(Space(14) & Format(rs!nSaldoIni, "#,#0.00"), 13) & COFF
        lnSaldo = rs!nSaldoIni
        nLin = nLin + 1
    End If
    lsMovNro = rs!cMovnro
        
    lnImporte = rs!nMovImporte
    lnTotD = lnTotD + rs!debe
    lnTotH = lnTotH + rs!haber
    lnSaldo = lnSaldo + lnImporte
    lsGlosa = rs!cMovDesc
    
    Linea lsTexto, Con & " " & Mid(lsMovNro, 1, 8) & "-" & Mid(lsMovNro, 9, 6) & "-" & Right(lsMovNro, 4) & ImpreFormat(lsGlosa, 40) & " " _
            & ImpreFormat(rs!Documento, 17) & " " _
            & ImpreFormat(rs!debe, 14, 2, True) & ImpreFormat(rs!haber, 14, 2, True) & " " _
            & ImpreFormat(lnSaldo, 10, 2, True) & COFF

    nLin = nLin + 1
    RaiseEvent BarraProgress(rs.Bookmark, "Generando Flujo de Bancos", Trim(rs!cPersNombre) & " [" & rs!cCtaIfCod + "]", "", vbBlue)
    lsCtaIFCod = rs!cCtaIfCod
    lsPersCod = rs!cPersCod
    rs.MoveNext
Loop
If lnTotD > 0 Or lnTotH > 0 Then
    Linea lsTexto, Con & " ------------------------------------------------------------------------------------------------------------------------------------"
    nLin = nLin + 1
    Linea lsTexto, BON & Space(80) & "TOTALES  " & ImpreFormat(lnTotD, 14, 2, True) & ImpreFormat(lnTotH, 14, 2, True) & BOFF & COFF, 2
    nLin = nLin + 2
End If
Linea lsTexto, Con & " ====================================================================================================================================" & COFF, 2
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
   Linea lsTexto, Cabecera(" FLUJO  DE  ENTIDADES  FINANCIERAS ", P, psSimbolo, pnColPage)
   Linea lsTexto, Centra("( DEL " & pdDesde & " AL " & pdHasta & ")", pnColPage) & BOFF
   Linea lsTexto, PrnSet("C+")
   nLin = 11
   Linea lsTexto, " ===================================================================================================================================="
   Linea lsTexto, " Número de             CONCEPTO                                       DOCUMENTO                        MOVIMIENTO             SALDOS "
   Linea lsTexto, " Movimiento                                                          Tipo  Número                 DEBE           HABER               "
   Linea lsTexto, " ------------------------------------------------------------------------------------------------------------------------------------" & PrnSet("B-")
End If
End Function
Public Function ImprimeSaldosCtaIf(ByVal pnColPage As Integer, ByVal pnLinPage As Integer, ByVal psOpeCod As String, _
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
lsSimbolo = IIf(Mid(psOpeCod, 3, 1) = gMonedaNacional, "S/.", "$.")
Set rs = oCtaIf.GetRepSaldosCtaIf(psOpeCod, pdFecha)
nLin = pnLinPage
If Not rs.EOF And Not rs.BOF Then
    RaiseEvent BarraShow(rs.RecordCount)
    Do While Not rs.EOF
       If Len(rs!cCtaIfCod) = 15 Then
          Linea lsTexto, ""
          nLin = nLin + 1
       End If
       RaiseEvent BarraProgress(rs.Bookmark, "Generando Saldos de Cuentas", "", "", vbBlue)
       CabeceraRepoSaldo lsTexto, nLin, P, pnLinPage, pnColPage, lsSimbolo, pdFecha, psOpeCod
       Linea lsTexto, Con & " " & Mid(rs!cCtaIfCod & Space(20), 1, 25) & "   " & Mid(rs!cCtaIFDesc & Space(70), 1, 55) & " " _
              & Right(Space(18) & Format(rs!debe, "#,#0.00"), 18) & " " & Right(Space(10) & Format(rs!haber, "#,#0.00"), 18) & " " & COFF
       nLin = nLin + 1
       rs.MoveNext
    Loop
    Linea lsTexto, Con & " ====================================================================================================================================" & COFF, 2
End If
rs.Close: Set rs = Nothing
RaiseEvent BarraClose
ImprimeSaldosCtaIf = lsTexto
End Function
Private Sub CabeceraRepoSaldo(ByRef lsTexto As String, ByRef nLin As Integer, ByRef P As Integer, _
                        ByVal pnLinPage As Integer, pnColPage As Integer, ByVal psSimbolo As String, ByVal pdHasta As Date, ByVal psOpeCod As String)
If nLin > pnLinPage - 6 Then
   Linea lsTexto, Cabecera(" SALDOS  DE  ENTIDADES  FINANCIERAS ", P, psSimbolo, pnColPage)
   Linea lsTexto, Centra("( AL " & pdHasta & " )", pnColPage) & BOFF
   Linea lsTexto, Con
   nLin = 11
   Linea lsTexto, " ============================================================================================================================"
   Linea lsTexto, "   ENTIDAD FINANCIERA                                                                                 S A L D O S   M. " & IIf(Mid(psOpeCod, 3, 1) = gMonedaNacional, "N.", "E.")
   Linea lsTexto, "                                                                                                    DEBE              HABER  "
   Linea lsTexto, " ----------------------------------------------------------------------------------------------------------------------------"
End If
End Sub

