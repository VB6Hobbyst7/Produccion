Attribute VB_Name = "gFunLogistica"
Option Explicit

'/*SUBASTA*/
Global Const gnSubIniProceso = 581001
Global Const gnSubListados = 581002
Global Const gnSubVentaFac = 581003
Global Const gnSubVentaBol = 581004
Global Const gnSubRegBilletaje = 581005
Global Const gnSubCuadreCaja = 581006
Global Const gnSubCierreProceso = 581007

Global Const gnSubProvMensual = 581101

'/*DEPRECIACION DE ACTIVO FIJO*/
Global Const gnDepAF = 581201
Global Const gnDepAjusteAF = 581202
Global Const gnDepDifAjusteAF = 581203
Global Const gnTransAF = 581204
Global Const gnDepTributAF = 581210 '*** PEAC 20120612
Global Const gnAsignaAF = 581205
Global Const gnAjusteVidaUtilAF = 581207 'EJVG20130701
Global Const gnDeterioroAF = 581211 'EJVG20130701

Global Const gnBajaAF = 581299

Global Const gnTransBND = 581301
Global Const gnAsignaBND = 581302
Global Const gnBajaBND = 581399

'/*ALMACEN*/
Global Const gnAlmaReqAreaReg = 591001
Global Const gnAlmaReqAreaMant = 591002
Global Const gnAlmaReqAreaExt = 591003
Global Const gnAlmaReqAreaRechPar = 591004

Global Const gnAlmaIngXCompras = 591101
Global Const gnAlmaIngXComprasConfirma = 591102
Global Const gnAlmaIngXTransferencia = 591103
Global Const gnAlmaIngXDevAreasGaran = 591104
Global Const gnAlmaIngXProvGaranRepa = 591105
Global Const gnAlmaIngXProvDemosOtros = 591106
Global Const gnAlmaIngXDacionPago = 591107
Global Const gnAlmaIngXEmbargo = 591108
Global Const gnAlmaIngXAdjudicacion = 591109
Global Const gnAlmaIngXOtrosMotivos = 591110

Global Const gnAlmaSalXAtencion = 591201
Global Const gnAlmaSalXTransferenciaOrigen = 591202
Global Const gnAlmaSalXProvGarantRepa = 591203
Global Const gnAlmaSalXAreasDevGaranRepa = 591204
Global Const gnAlmaSalXProvDevCompras = 591205
Global Const gnAlmaSalXProvDemosOtros = 591206
Global Const gnAlmaSalXDevolEmbargo = 591207
Global Const gnAlmaSalXOtrosMotivos = 591208
Global Const gnAlmaSalXAjuste = 591209

Global Const gnAlmaMantXIngreso = 591301
Global Const gnAlmaMantXSalida = 591302

Global Const gnAlmaExtornoXIngreso = 591401
Global Const gnAlmaExtornoXSalida = 591402
Global Const gnAlmaExtornoXConfirmacionIng = 591403
'EJVG20131015 ***
Global Const gnAlmaContratoRegistroMN = 501215 'PASI20140110 ERS0772014
Global Const gnAlmaContratoRegistroME = 502215 'PASI20140110 ERS0772014
Global Const gnAlmaActaConformidadMN = 591601
Global Const gnAlmaActaConformidadLibreMN = 591602
Global Const gnAlmaActaConformidadME = 592601
Global Const gnAlmaActaConformidadLibreME = 592602
Global Const gnAlmaActaConformidadExtornoMN = 591605
Global Const gnAlmaActaConformidadExtornoME = 592605
Global Const gnAlmaComprobanteRegistroMN = 591701 'PASI20140923 inteneto Cambio valor x 501216
Global Const gnAlmaComprobanteLibreRegistroMN = 591702 'PASI20140923 ERS0772014
Global Const gnAlmaComprobanteExtornoMN = 591703 'PASI20140923 ERS0772014
Global Const gnAlmaComprobanteLibreExtornoMN = 591704 'PASI20140923 ERS0772014
Global Const gnAlmaComprobanteRegistroME = 592701 'PASI20140923 inteneto Cambio valor x 502216
Global Const gnAlmaComprobanteLibreRegistroME = 592702 'PASI20140923 ERS0772014
Global Const gnAlmaComprobanteExtornoME = 592703 'PASI20140923 ERS0772014
Global Const gnAlmaComprobanteLibreExtornoME = 592704 'PASI20140923 ERS0772014
'END EJVG *******

Public sObtTraNro As String

Public Enum LogTipoOC
    gLogOCompraDirecta = 130
    gLogOServicioDirecta = 132
    gLogOCompraProceso = 133
    gLogOServicioProceso = 134
End Enum
'EJVG20131015 ***
Public Enum LogTipoActaConformidad
    gActaRecepcionBienes = 148
    gActaConformidadServicio = 149
End Enum
Public Enum LogTipoPagoComprobante
    gPagoCuentaCMAC = 1
    gPagoTransferencia = 2
    gPagoCheque = 3
End Enum
'END EJVG *******

Global Const gLogOCDirecta = "D"
Global Const gLogOCProceso = "P"

Global Const gsLogistica = "023"

'Devuelve un String con un Nro de Mov
Public Function GeneraMovCorre(ByVal psMovNro As String) As String
    Dim sCorre As String
    sCorre = FillNum(Str(Val(Mid(sObtTraNro, 21, 2)) + 1), 2, "0")
    GeneraMovCorre = Left(psMovNro, 19) & sCorre & Right(psMovNro, 4)
End Function

Public Function GeneraActualizacion(ByVal psFecSis As Date, ByVal psCmac As String, _
ByVal psAgencia As String, ByVal psUsuario As String) As String
    GeneraActualizacion = Format(psFecSis, "yyyymmdd") & Format(Time, "hhmmss") & psCmac & Right(psAgencia, 2) & "00" & psUsuario
End Function

Public Function GeneraCotiza(ByVal psSelNro As String, ByVal pnCorrela As Integer) As String
    GeneraCotiza = Left(psSelNro, 19) & FillNum(Str(pnCorrela), 2, "0") & Right(psSelNro, 4)
End Function

'EJVG20131023 ***
Public Sub ImprimeActaConformidadPDF(ByVal pnMovNro As Long)
    Dim oLog As New DLogGeneral
    Dim odoc As New cPDF
    Dim R As New ADODB.Recordset
    
    Set R = oLog.ActaConformidadxImpresion(pnMovNro)
    If R.EOF Then Exit Sub
    
    odoc.Author = gsCodUser
    odoc.Creator = "SICMACT - Administrativo"
    odoc.Producer = gsNomCmac
    odoc.Subject = "ACTA DE CONFORMIDAD Nº " & R!cDocNro
    odoc.Title = "ACTA DE CONFORMIDAD Nº " & R!cDocNro
    
    If Not odoc.PDFCreate(App.path & "\Spooler\ActaConformidad_" & R!cDocNro & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If

    odoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
    odoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
    odoc.LoadImageFromFile App.path & "\logo_cmacmaynas.bmp", "Logo"
    odoc.NewPage A4_Vertical
    
    odoc.WImage 75, 40, 35, 105, "Logo"
    'odoc.WTextBox 63, 40, 15, 500, IIf(R!nDocTpo = LogTipoActaConformidad.gActaRecepcionBienes, "ACTA DE RECEPCION DE BIENES", "ACTA DE CONFORMIDAD DE SERVICIOS"), "F2", 12, hCenter
    odoc.WTextBox 63, 40, 15, 500, IIf(R!nDocTpo = LogTipoActaConformidad.gActaRecepcionBienes, "ACTA DE CONFORMIDAD DE PAGO", "ACTA DE CONFORMIDAD DE PAGO"), "F2", 12, hCenter 'AGREGADO POR VAPA20170415
    odoc.WTextBox 83, 40, 15, 500, R!cDocNro, "F2", 12, hCenter
    odoc.WTextBox 100, 40, 705, 520, R!cAgeDistrito & ", " & Format(Day(R!dDocFecha), "00") & " de " & Format(R!dDocFecha, "mmmm") & " del " & Year(R!dDocFecha), "F1", 10, AlignH.hRight, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox 150, 40, 705, 250, "Área que requiere el Bien: ", "F2", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox 150, 180, 705, 380, R!cAreaAgeDesc, "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox 165, 180, 705, 380, R!cSubAreaDesc, "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox 200, 40, 705, 250, "Bien enviado por: ", "F2", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox 200, 180, 705, 380, R!cProveedorNombre, "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox 235, 40, 705, 520, "Descripción: ", "F2", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox 250, 40, 705, 520, R!cDescripcion, "F1", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox 400, 40, 705, 520, "Observación: ", "F2", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox 415, 40, 705, 520, R!cObservacion, "F1", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
    If R!cIFiCtaCod <> "" And R!cIFiNombre <> "" Then
        odoc.WTextBox 530, 40, 705, 150, "Institución:", "F2", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox 530, 90, 705, 470, R!cIFiNombre, "F1", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox 540, 40, 705, 150, "CTA N°:", "F2", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox 540, 90, 705, 470, R!cIFiCtaCod, "F1", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
    End If
    odoc.WTextBox 580, 40, 705, 250, "Firma de Conformidad: ", "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox 600, 40, 705, 250, "Solicitante ", "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox 660, 40, 705, 250, "Nombre: ", "F2", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox 660, 90, 705, 470, R!cUsuarioNombre, "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox 690, 40, 705, 520, "DNI N°:", "F2", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox 690, 90, 705, 470, R!cUsuarioDNI, "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    
    odoc.PDFClose
    odoc.Show
    
    Set R = Nothing
    Set oLog = Nothing
    Set odoc = Nothing
End Sub
'END EJVG *******
'PASI20140929 ERS0772014
Public Sub ImprimeActaConformidadPDFNew(ByVal pnMovNro As Long, ByVal pnMoneda As Integer, ByVal fnTpo As Integer)
    Dim oLog As New DLogGeneral
    Dim odoc As New cPDF
    Dim nEspa As Integer
    Dim R As New ADODB.Recordset
    Dim Rdet As New ADODB.Recordset
    
    Set R = oLog.ActaConformidad_ERS0772014_xImpresion(pnMovNro)
    
    If fnTpo = LogTipoDocOrigenComprobante.OrdenServicio _
    Or fnTpo = LogTipoDocOrigenComprobante.OrdenCompra _
    Or fnTpo = LogTipoDocOrigenComprobante.CompraLibre _
    Or fnTpo = LogTipoDocOrigenComprobante.Serviciolibre Then
        Set Rdet = oLog.ActaConformidadOrdenDet_ERS0772014_xImpresion(pnMovNro, pnMoneda)
    Else
        Set Rdet = oLog.ActaConformidadContDet_ERS0772014_xImpresion(pnMovNro, pnMoneda)
    End If
    If R.EOF Then Exit Sub
    If Rdet.EOF Then Exit Sub
    
    odoc.Author = gsCodUser
    odoc.Creator = "SICMACT - Administrativo"
    odoc.Producer = gsNomCmac
    odoc.Subject = "ACTA DE CONFORMIDAD Nº " & R!cDocNro
    odoc.Title = "ACTA DE CONFORMIDAD Nº " & R!cDocNro
    
    If Not odoc.PDFCreate(App.path & "\Spooler\ActaConformidad_" & R!cDocNro & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If

    odoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding 'PASI20151112 Cambio de Arial a Tahoma
    odoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding 'PASI20151112 Cambio de Arial a Tahoma
    odoc.LoadImageFromFile App.path & "\logo_cmacmaynas.bmp", "Logo"
    odoc.NewPage A4_Vertical
    
    Dim nposicion As Integer
    nposicion = 63
    
    odoc.WImage 75, 40, 35, 105, "Logo"
    'oDoc.WTextBox nposicion, 40, 15, 500, IIf(R!nDocTpo = LogTipoActaConformidad.gActaRecepcionBienes, "ACTA DE RECEPCION DE BIENES", "ACTA DE CONFORMIDAD DE SERVICIOS"), "F2", 12, hCenter 'COMENTADO POR VAPA20170415
    odoc.WTextBox nposicion, 40, 15, 500, IIf(R!nDocTpo = LogTipoActaConformidad.gActaRecepcionBienes, "ACTA DE CONFORMIDAD DE PAGO", "ACTA DE CONFORMIDAD DE PAGO"), "F2", 12, hCenter 'AGREGO VAPA20170415
    odoc.WTextBox nposicion + 20, 40, 15, 500, R!cDocNro, "F2", 12, hCenter
    odoc.WTextBox nposicion + 40, 40, 705, 520, R!cAgeDistrito & ", " & Format(Day(R!dDocFecha), "00") & " de " & Format(R!dDocFecha, "mmmm") & " del " & Year(R!dDocFecha), "F1", 10, AlignH.hRight, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox nposicion + 90, 40, 705, 250, "Área que requiere el Bien: ", "F2", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox nposicion + 90, 180, 705, 380, R!cAreaAgeDesc, "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox nposicion + 100, 180, 705, 380, R!cSubAreaDesc, "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox nposicion + 140, 40, 705, 250, "Bien enviado por: ", "F2", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox nposicion + 140, 180, 705, 380, R!cProveedorNombre, "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox nposicion + 175, 40, 705, 520, "Descripción: ", "F2", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox nposicion + 190, 40, 705, 520, R!cDescripcion, "F1", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox nposicion + 215, 40, 705, 520, "Datos de Comprobante: ", "F2", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox nposicion + 230, 40, 705, 50, "Tipo Doc:", "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox nposicion + 230, 90, 705, 150, R!TipoDoc, "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox nposicion + 230, 220, 705, 50, "Nro Doc:", "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox nposicion + 230, 270, 705, 100, R!NroDoc, "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox nposicion + 230, 370, 705, 80, "Fecha Emisión:", "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox nposicion + 230, 450, 705, 100, Format(R!FechaEmision, "dd/mm/yyyy"), "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox nposicion + 245, 40, 705, 100, "Doc. Origen: ", "F2", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox nposicion + 265, 40, 705, 300, R!DocOrigen, "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    'Cabecera del Comprobante
    
    nposicion = 340
        odoc.WTextBox nposicion, 40, 15, 700, String(134, "_"), "F2", 7, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox nposicion + 10, 40, 15, 250, "Descripción", "F2", 7, AlignH.hCenter, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox nposicion + 10, 290, 15, 5, "|", "F2", 7, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox nposicion + 10, 295, 15, 60, "Unidad", "F2", 7, AlignH.hCenter, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox nposicion + 10, 355, 15, 5, "|", "F2", 7, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox nposicion + 10, 360, 15, 60, "Solicitado", "F2", 7, AlignH.hCenter, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox nposicion + 10, 420, 15, 5, "|", "F2", 7, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox nposicion + 10, 425, 15, 60, "P. Unitario", "F2", 7, AlignH.hCenter, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox nposicion + 10, 485, 15, 5, "|", "F2", 7, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox nposicion + 10, 490, 15, 60, "Total", "F2", 7, AlignH.hCenter, vTop, vbBlack, 0, vbBlack
        'odoc.WTextBox nposicion + 10, 550, 15, 5, "|", "F2", 7, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox nposicion + 12, 40, 15, 700, String(134, "_"), "F2", 7, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        
        Dim I, j, nlineas, varpos, lineaBasePdf As Integer
        Dim lnBaseDetalle As Integer
        Dim varDesc As String
        Dim varTope As Integer
        I = 1
        varTope = 800
        lnBaseDetalle = 365
        varpos = 1
        lineaBasePdf = 365
        
        Do While Not Rdet.EOF
            If I = 1 Then
                odoc.WTextBox lnBaseDetalle, 295, 15, 60, Rdet!unidad, "F1", 7, AlignH.hCenter, vTop, vbBlack, 0, vbBlack 'Unidad
                'oDoc.WTextBox lnBaseDetalle, 360, 15, 60, Format(Rdet!solicitado, "#,#"), "F1", 7, AlignH.hCenter, vTop, vbBlack, 0, vbBlack 'Solicitado
                odoc.WTextBox lnBaseDetalle, 360, 15, 60, Rdet!solicitado, "F1", 7, AlignH.hCenter, vTop, vbBlack, 0, vbBlack 'Solicitado 'PASI20151127 **
                odoc.WTextBox lnBaseDetalle, 425, 15, 60, Format(Rdet!PrecUnit, "#,#0.000"), "F1", 7, AlignH.hRight, vTop, vbBlack, 0, vbBlack 'Prec Unit
                odoc.WTextBox lnBaseDetalle, 490, 15, 60, Format(Rdet!Total, "#,#0.00"), "F1", 7, AlignH.hRight, vTop, vbBlack, 0, vbBlack 'Total
                varDesc = JustificaTextoCadenaPASI(Rdet!Descripcion, 40, 1)
                nlineas = Round((Len(varDesc) / 61) + 0.4)
                odoc.WTextBox lnBaseDetalle, 40, 15, 240, varDesc, "F1", 7, AlignH.hjustify, vTop, vbBlack, 0, vbBlack 'Descripcion
                lnBaseDetalle = lnBaseDetalle + nlineas * 5
            Else
                varDesc = JustificaTextoCadenaPASI(Rdet!Descripcion, 40, 1)
                nlineas = Round((Len(varDesc) / 61) + 0.4)
                lineaBasePdf = lnBaseDetalle + nlineas * 5
                If lineaBasePdf >= varTope Then
                    odoc.NewPage A4_Vertical
                   lnBaseDetalle = 43
                End If
                odoc.WTextBox lnBaseDetalle, 295, 15, 60, Rdet!unidad, "F1", 7, AlignH.hCenter, vTop, vbBlack, 0, vbBlack 'Unidad
                'oDoc.WTextBox lnBaseDetalle, 360, 15, 60, Format(Rdet!solicitado, "#,#"), "F1", 7, AlignH.hCenter, vTop, vbBlack, 0, vbBlack 'Solicitado
                odoc.WTextBox lnBaseDetalle, 360, 15, 60, Rdet!solicitado, "F1", 7, AlignH.hCenter, vTop, vbBlack, 0, vbBlack 'Solicitado 'PASI20151127 **
                odoc.WTextBox lnBaseDetalle, 425, 15, 60, Format(Rdet!PrecUnit, "#,#0.000"), "F1", 7, AlignH.hRight, vTop, vbBlack, 0, vbBlack 'Prec Unit
                odoc.WTextBox lnBaseDetalle, 490, 15, 60, Format(Rdet!Total, "#,#0.00"), "F1", 7, AlignH.hRight, vTop, vbBlack, 0, vbBlack 'Total
                odoc.WTextBox lnBaseDetalle, 40, 15, 240, varDesc, "F1", 7, AlignH.hjustify, vTop, vbBlack, 0, vbBlack    'Descripcion
                lnBaseDetalle = lnBaseDetalle + nlineas * 5
            End If
            I = I + 1
            lnBaseDetalle = lnBaseDetalle + 20
            Rdet.MoveNext
        Loop
        lnBaseDetalle = lnBaseDetalle - 5
        odoc.WTextBox lnBaseDetalle, 40, 15, 700, String(134, "_"), "F2", 7, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        Dim q As Integer
            For q = 0 To 8
                lnBaseDetalle = lnBaseDetalle + 10
                If lnBaseDetalle >= varTope Then
                    odoc.NewPage A4_Vertical
                    lnBaseDetalle = 43
                End If
            Next
        
        odoc.WTextBox lnBaseDetalle, 40, 705, 520, "Observación: ", "F2", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox lnBaseDetalle + 10, 40, 705, 520, R!cObservacion, "F1", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack

        lnBaseDetalle = lnBaseDetalle + 20
        If lnBaseDetalle >= varTope Then
            odoc.NewPage A4_Vertical
            lnBaseDetalle = 43
        End If
        
        If R!cGuia <> "N/A" Then
            odoc.WTextBox lnBaseDetalle + 10, 40, 705, 520, "Guia:", "F2", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
            odoc.WTextBox lnBaseDetalle + 10, 90, 705, 520, R!cGuia, "F1", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
        End If

        lnBaseDetalle = lnBaseDetalle + 100
        If lnBaseDetalle >= varTope Then
            odoc.NewPage A4_Vertical
            lnBaseDetalle = 43
        End If
        
    If R!cIFiCta <> "" And R!cIFiNombre <> "" Then
    Dim ificta As String
        odoc.WTextBox lnBaseDetalle, 40, 705, 100, "Institución:", "F2", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox lnBaseDetalle, 100, 705, 500, R!cIFiNombre, "F1", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
         lnBaseDetalle = lnBaseDetalle + 10
        If lnBaseDetalle >= varTope Then
            odoc.NewPage A4_Vertical
            lnBaseDetalle = 43
        End If
        odoc.WTextBox lnBaseDetalle, 40, 705, 60, "CTA N°:", "F2", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox lnBaseDetalle, 100, 705, 120, R!cIFiCta, "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    End If
    '********PASI20151112 ERS0472015*******
    lnBaseDetalle = lnBaseDetalle + 10
    If lnBaseDetalle >= varTope Then
        odoc.NewPage A4_Vertical
        lnBaseDetalle = 43
    End If
    If R!CCI <> "" Then
        odoc.WTextBox lnBaseDetalle, 40, 705, 100, "CCI:", "F2", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox lnBaseDetalle, 100, 705, 500, R!CCI, "F1", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
    End If
    lnBaseDetalle = lnBaseDetalle + 10
    If lnBaseDetalle >= varTope Then
        odoc.NewPage A4_Vertical
        lnBaseDetalle = 43
    End If
    If R!BancoDetrac <> "" Then
        odoc.WTextBox lnBaseDetalle, 40, 705, 100, "Bco. Detrac:", "F2", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox lnBaseDetalle, 100, 705, 500, R!BancoDetrac, "F1", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
    End If
    lnBaseDetalle = lnBaseDetalle + 10
    If lnBaseDetalle >= varTope Then
        odoc.NewPage A4_Vertical
        lnBaseDetalle = 43
    End If
    If R!cCtaDetrac <> "" Then
        odoc.WTextBox lnBaseDetalle, 40, 705, 100, "Cta. Detrac:", "F2", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
        odoc.WTextBox lnBaseDetalle, 100, 705, 500, R!cCtaDetrac, "F1", 10, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
    End If
    lnBaseDetalle = lnBaseDetalle + 10
    '*******end PASI***********************
    
    If lnBaseDetalle + 80 >= varTope Then
        odoc.NewPage A4_Vertical
        lnBaseDetalle = 43
    End If
    lnBaseDetalle = lnBaseDetalle + 80
     odoc.WTextBox lnBaseDetalle, 40, 705, 250, "Firma de Conformidad: ", "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
     lnBaseDetalle = lnBaseDetalle + 10
     If lnBaseDetalle >= varTope Then
        odoc.NewPage A4_Vertical
        lnBaseDetalle = 43
    End If
     odoc.WTextBox lnBaseDetalle, 40, 705, 250, "Solicitante ", "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
     If lnBaseDetalle + 60 >= varTope Then
        odoc.NewPage A4_Vertical
        lnBaseDetalle = 43
    End If
    lnBaseDetalle = lnBaseDetalle + 60
    odoc.WTextBox lnBaseDetalle, 40, 705, 250, "Nombre: ", "F2", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox lnBaseDetalle, 90, 705, 470, R!cUsuarioNombre, "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    lnBaseDetalle = lnBaseDetalle + 10
    If lnBaseDetalle >= varTope Then
        odoc.NewPage A4_Vertical
        lnBaseDetalle = 43
    End If
    odoc.WTextBox lnBaseDetalle, 40, 705, 520, "DNI N°:", "F2", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    odoc.WTextBox lnBaseDetalle, 90, 705, 470, R!cUsuarioDNI, "F1", 10, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    
    odoc.PDFClose
    odoc.Show
    
    Set R = Nothing
    Set Rdet = Nothing
    Set oLog = Nothing
    Set odoc = Nothing
End Sub
'end PASI

