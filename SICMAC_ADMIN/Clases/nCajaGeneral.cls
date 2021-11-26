VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nCajaGeneral"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsServerCom As String
Dim vsServerPers As String
Dim vsConexion As String
Public Enum MotivosDepRet
    gMotDepositos = 1
    gMotRetiros = 2
End Enum

Dim gOpeMECompraAInst As String
Dim gOpeMEVentaAInst As String

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

Dim oIni As ClasIni
Set oIni = New ClasIni
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
End Sub
Public Function GetNroCodOp() As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oconect As DConecta

Set oconect = New DConecta
If oconect.AbreConexion = False Then Exit Function
Set rs = New ADODB.Recordset
GetNroCodOp = ""
sql = "Select Isnull(Max(cRangoCod),'0') as MaxCodRango From OpCajaGen"
Set rs = oconect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetNroCodOp = Format(Val(rs!maxCodRango) + 1, "0000")
End If
rs.Close
Set rs = Nothing
oconect.CierraConexion
Set oconect = Nothing
End Function
Public Function GetEstadosOp(Optional pbFiltro As Boolean, Optional pnbEstOpDesde As CGEstadosOp = PorEmitir, Optional pnbEstOpHasta As CGEstadosOp = Extraviadas) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oconect As DConecta

Set rs = New ADODB.Recordset
Set oconect = New DConecta

sql = " Select  cConsDescripcion, nConsValor " _
    & " From    Constante " _
    & " Where   nCONSCOD= '" & gCGEstadosOp & "' and nCONSCOD<>nConsValor "

If oconect.AbreConexion = False Then Exit Function
Set rs = oconect.CargaRecordSet(sql)
If pbFiltro Then
     rs.Filter = "nConsValor >='" & pnbEstOpDesde & "'  AND nConsValor<= '" & pnbEstOpHasta & "'"
End If
Set GetEstadosOp = rs
GetEstadosOp.MoveFirst
oconect.CierraConexion
Set oconect = Nothing
End Function
Public Function GrabaRangoOpCajaGeneral(ByVal psFormatoFecha As String, ByVal psRangoCod As String, ByVal pdFechaIng As Date, ByVal pnRangoEmiIni As Long, ByVal pnRangoEmiFin As Long, ByVal psEstadoRango As CGEstadosOp, ByVal psUltActualizacion As String, ByVal pnMoneda As Moneda, Optional pbEjectBacth As Boolean = False) As Integer
Dim oCajaGen As DCajaGeneral
GrabaRangoOpCajaGeneral = 1
Set oCajaGen = New DCajaGeneral
oCajaGen.Inicio psFormatoFecha
oCajaGen.InsertaOpCajaGen psRangoCod, pdFechaIng, pnRangoEmiIni, pnRangoEmiFin, psEstadoRango, psUltActualizacion, pnMoneda, pbEjectBacth
If pbEjectBacth Then
    GrabaRangoOpCajaGeneral = oCajaGen.EjecutaBatch
Else
    GrabaRangoOpCajaGeneral = 0
End If
Set oCajaGen = Nothing
End Function
Public Function GrabaOpDetCajaGeneral(ByVal psRangoCod As String, ByVal psDescripcion As String, ByVal pnRangoEmiIni As Long, ByVal pnRangoEmiFin As Long, ByVal psEstadoRango As CGEstadosOp, ByVal psUltActualizacion As String, Optional pbEjectBacth As Boolean = False) As Integer
Dim oCajaGen As DCajaGeneral
GrabaOpDetCajaGeneral = 1
Set oCajaGen = New DCajaGeneral
oCajaGen.InsertaOpDetCajaGen psRangoCod, pnRangoEmiIni, pnRangoEmiFin, psEstadoRango, psDescripcion, psUltActualizacion, pbEjectBacth
If pbEjectBacth Then
    GrabaOpDetCajaGeneral = oCajaGen.EjecutaBatch
Else
    GrabaOpDetCajaGeneral = 0
End If
Set oCajaGen = Nothing
End Function
Public Function EliminaOpDetCaja(ByVal psRangoCod As String, ByVal pnRangoIni As Long, ByVal pnRangoFin As Long, Optional pbEjectBacth As Boolean = False) As Integer
Dim oCajaGen As DCajaGeneral
EliminaOpDetCaja = 1
Set oCajaGen = New DCajaGeneral
oCajaGen.EliminaOpDetCajaGen psRangoCod, pnRangoIni, pnRangoFin, pbEjectBacth
If pbEjectBacth Then
    EliminaOpDetCaja = oCajaGen.EjecutaBatch
Else
    EliminaOpDetCaja = 0
End If
Set oCajaGen = Nothing
End Function


Public Function UpdateOpCajaGeneral(ByVal psRangoCod As String, ByVal pnRangoEmiIni As Long, ByVal pnRangoEmiFin As Long, ByVal psUltActualizacion As String, Optional pbEjectBacth As Boolean = False) As Integer
Dim oCajaGen As DCajaGeneral
UpdateOpCajaGeneral = 1
Set oCajaGen = New DCajaGeneral
oCajaGen.ActualizaOpCajaGen psRangoCod, pnRangoEmiIni, pnRangoEmiFin, psUltActualizacion, pbEjectBacth
If pbEjectBacth Then
    UpdateOpCajaGeneral = oCajaGen.EjecutaBatch
Else
    UpdateOpCajaGeneral = 0
End If
Set oCajaGen = Nothing
End Function

Public Function EliminaOpCajaGeneral(ByVal psRangoCod As String, Optional pbEjectBacth As Boolean = False) As Integer
Dim oCajaGen As DCajaGeneral
EliminaOpCajaGeneral = 1
Set oCajaGen = New DCajaGeneral
oCajaGen.EliminaOpCajaGen psRangoCod, pbEjectBacth
If pbEjectBacth Then
    EliminaOpCajaGeneral = oCajaGen.EjecutaBatch
Else
    EliminaOpCajaGeneral = 0
End If
Set oCajaGen = Nothing
End Function

Public Function GetOpCajaGeneral() As ADODB.Recordset
Dim sql As String
Dim oconect As DConecta
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
Set oconect = New DConecta
If oconect.AbreConexion = False Then Exit Function

sql = "Select   cRangoCod, dFechaIng, nRangoEmiIni, nRangoEmiFin, " _
    & "         cEstadoRango, C.cConsDescripcion,  " _
    & "         Op.cMoneda, C1.cConsDescripcion as cConsDescMon, cUltActualizacion " _
    & " From    OPCajaGen op JOIN Constante C on C.nConsValor= Op.cEstadoRango " _
    & "         JOIN Constante C1 on C1.nConsValor= Op.cMoneda " _
    & " where   C.nCONSCOD= '" & gCGEstadosOp & "' and C1.nCONSCOD='" & gMoneda & "' ORDER BY OP.cRangoCod "
    
Set rs = oconect.CargaRecordSet(sql)
Set GetOpCajaGeneral = rs
oconect.CierraConexion
Set oconect = Nothing
End Function
Public Function GetOpDetCajaGeneral(ByVal psRangoCod As String) As ADODB.Recordset
Dim sql As String
Dim oconect As DConecta
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
Set oconect = New DConecta
If oconect.AbreConexion = False Then Exit Function

sql = "Select cRangoCod, nRangoIni,nRangoFin, cEstadoDet, cConsDescripcion, cDescDet" _
    & " From OPCajaGenDet op JOIN Constante C on C.nConsValor= Op.cEstadoDet " _
    & " where C.nCONSCOD= '" & gCGEstadosOp & "' and OP.cRangoCod ='" & psRangoCod & "'"
Set rs = oconect.CargaRecordSet(sql)
Set GetOpDetCajaGeneral = rs
oconect.CierraConexion
Set oconect = Nothing
End Function
Public Function VerificaOPRango(ByVal psDocNro As String, ByVal psMoneda As Moneda) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oconect As DConecta
Set oconect = New DConecta
Set rs = New ADODB.Recordset

If oconect.AbreConexion = False Then Exit Function

'PRIMERO VERIFICAMOS SI EL DOCUMENTO SE ENCUENTRA DENTRO DEL RANGO ESTABLECIDO
sql = "     SELECT  OPG.cRangoCod,nRangoEmiIni,nRangoEmiFin,cMoneda," _
        & "         nRangoIni As RangoInIDet, nRangoFin As RangoFinDet, cEstadoDet " _
        & " From    OPCajaGen OPG LEFT JOIN OPCajaGenDet OPD ON OPD.cRangoCod= OPG.cRangoCod " _
        & " WHERE   OPG.CMONEDA ='" & psMoneda & "' AND OPG.cEstadoRango='" & PorEmitir & "' " _
        & "         AND " & psDocNro & " BETWEEN OPG.nRangoEmiIni AND OPG.nRangoEmiFin "

Set rs = oconect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.EOF Then
    'si se encuentra la op dentro del rango de las ordenes emitidas
    'buscamos ahora dentro del rango de ordenes de pagos anuladas
    'es decir en la tabla de Ordenes de pagoDetalle
    If IsNull(rs!RangoIniDet) = False Then
        sql = " SELECT   OPG.cRangoCod,nRangoEmiIni,nRangoEmiFin,cMoneda," _
            & "         nRangoIni As RangoInIDet, nRangoFin As RangoFinDet, cEstadoDet " _
            & " From    OPCajaGen OPG LEFT JOIN OPCajaGenDet OPD ON OPD.cRangoCod= OPG.cRangoCod " _
            & " WHERE   OPG.CMONEDA ='" & psMoneda & "' AND cEstadoRango='" & PorEmitir & "' " _
            & "         AND " & psDocNro & " BETWEEN OPD.nRangoIni AND OPD.nRangoFin"
        
        Set rs = oconect.CargaRecordSet(sql)
        If Not rs.EOF And Not rs.BOF Then
            'si se encuentra dentro del rango de algunas anuladas entonces
            'cogemos el ultimo rango de las seleccionadas y le aumentamos 1
            'y volvemos a verificar el documento
            VerificaOPRango = Format(VerificaOPRango(rs!RangoFinDet + 1, psMoneda), String(8, "0"))
        Else
            VerificaOPRango = Format(psDocNro, String(8, "0"))
        End If
        
    Else
        VerificaOPRango = Format(psDocNro, String(8, "0"))
    End If
Else
    'POSTERIORMENTE VERIFICAMOS SACAMOS SI EXISTE NUEVO RANGO DE OP INGRESADAS LISTAS PARA EMITIR
    sql = " SELECT  OPG.cRangoCod,nRangoEmiIni,nRangoEmiFin,cMoneda," _
        & "         nRangoIni As RangoInIDet, nRangoFin As RangoFinDet, cEstadoDet " _
        & " From    OPCajaGen OPG LEFT JOIN OPCajaGenDet OPD ON OPD.cRangoCod= OPG.cRangoCod " _
        & " WHERE   OPG.CMONEDA ='" & psMoneda & "' AND cEstadoRango='" & PorEmitir & "' "
    Set rs = oconect.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        VerificaOPRango = Format(rs!nRangoEmiIni, String(8, "0"))
    Else
        'Debera ingresar un nuevo rango de ordenes de pago listas para emitir
        VerificaOPRango = ""
    End If
End If
rs.Close
Set rs = Nothing
oconect.CierraConexion
Set oconect = Nothing
End Function
Public Function GrabaHabEfectivo(ByVal psMovNroHab As String, _
                                ByVal psFormatoFecha As String, ByVal rsBillete As ADODB.Recordset, _
                                ByVal rsMoneda As ADODB.Recordset, _
                                ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                ByVal psCtaDebe As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                ByVal psAreaOrig As String, ByVal psAreaDest As String, _
                                ByVal psNroDocTransp As String, ByVal psFechaDocTransp As String, _
                                ByVal psNroDocSal As String, ByVal psFechaDocSal As String, _
                                ByVal psCodTransp As String, ByVal pnMontoTransp As Currency, _
                                Optional ByVal psMovNroRef As String = "", Optional ByVal psAreaCodCaja As String = "") As Integer
                                
            
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim oContF As NContFunciones
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New DMov
On Error GoTo GrabaHabEfectivoErr

Set oContF = New NContFunciones

GrabaHabEfectivo = 1
oDMov.Inicio psFormatoFecha
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNroHab, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNroHab)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1:
lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaDebe, psAreaDest, False)
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe + lsSubCta, pnImporte
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(psAreaDest, 4, 2), Mid(psAreaDest, 1, 3)
'guardamos el billetaje
If Not rsBillete Is Nothing Then
    If rsBillete.State = adStateOpen Then
        If Not rsBillete.EOF And Not rsBillete.BOF Then
            Do While Not rsBillete.EOF
                If rsBillete!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaHaber, psAreaOrig, False)
                    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber + lsSubCta, rsBillete!Monto * -1
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oDMov.InsertaMovObjEfectivo lnMovNro, Format(lnMovItem, "#0"), Format(lnMovOrden, "#0"), rsBillete!cEfectivoCod
                End If
                rsBillete.MoveNext
            Loop
        End If
        rsBillete.Close: Set rsBillete = Nothing
    End If
End If
If Not rsMoneda Is Nothing Then
    If rsMoneda.State = adStateOpen Then
        If Not rsMoneda.EOF And Not rsMoneda.BOF Then
            Do While Not rsMoneda.EOF
                If rsMoneda!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaHaber, psAreaOrig, False)
                    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber + lsSubCta, rsMoneda!Monto * -1
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oDMov.InsertaMovObjEfectivo lnMovNro, Format(lnMovItem, "#0"), Format(lnMovOrden, "#0"), rsMoneda!cEfectivoCod
                End If
                rsMoneda.MoveNext
            Loop
        End If
        rsMoneda.Close: Set rsMoneda = Nothing
    End If
End If
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(psAreaOrig, 4, 2), Mid(psAreaOrig, 1, 3)

If psNroDocSal <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocSalidaEfectivoBoveda, psNroDocSal, psFechaDocSal
End If
If psNroDocTransp <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocCompServTranspValores, psNroDocTransp, psFechaDocTransp
End If
If psCodTransp <> "" Then
    If psAreaCodCaja = "" Then
        oDMov.InsertaMovHabilitaciones lnMovNro, psCodTransp, Mid(psAreaOrig, 1, 3), Mid(psAreaOrig, 4, 2), Mid(psAreaDest, 1, 3), Mid(psAreaDest, 4, 2), pnMontoTransp, pnImporte
    Else
        oDMov.InsertaMovHabilitaciones lnMovNro, psCodTransp, Mid(psAreaOrig, 1, 3), Mid(psAreaOrig, 4, 2), Mid(psAreaCodCaja, 1, 3), Mid(psAreaCodCaja, 4, 2), pnMontoTransp, pnImporte
    End If
End If
oDMov.CommitTrans
lbTrans = False

GrabaHabEfectivo = 0
Set oDMov = Nothing
Set oContF = Nothing
Exit Function
GrabaHabEfectivoErr:
    If lbTrans Then
        oDMov.RollBackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaHabEfectivo", Err.Description & " Error Graba Habilitación de Efectivo"
End Function
Public Function GetDatosHabilitacion(ByVal psMovNro As String) As ADODB.Recordset
Dim sql As String
Dim oconect As DConecta
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
Set oconect = New DConecta
If oconect.AbreConexion = False Then Exit Function

sql = "SELECT   M.CMOVNRO, " _
    & "         MH.CAREAORIG + MH.CAGEORIG + '- ' + ISNULL(AG.CAGEDESCRIPCION,A.CAREADESCRIPCION ) AS ORIGEN, " _
    & "         MH.CAREADEST + MH.CAGEDEST + '- ' + ISNULL(AG1.CAGEDESCRIPCION,A1.CAREADESCRIPCION ) AS DESTINO, " _
    & "         P.CPERSCOD , ISNULL(ISNULL(P.CPERSNOMBRE, T.cDescTrans ),'') AS cNombreTransp, MC.NMOVIMPORTE , ISNULL(MH.nMontoTrans,0) AS nMontoTrans,  ISNULL(C.cConsDescripcion,'') AS cTipoTrans, " _
    & "         ISNULL((SELECT D.CDOCABREV + '-' + MD.CDOCNRO FROM MOVDOC MD JOIN DOCUMENTO D ON D.nDocTpo =MD.nDocTpo WHERE MD.nMOVNRO=M.nMOVNRO AND MD.nDocTpo ='" & TpoDocCompServTranspValores & "' ),'')   AS DocCompTransp, " _
    & "         ISNULL((SELECT D.CDOCABREV + '-' + MD.CDOCNRO FROM MOVDOC MD JOIN DOCUMENTO D ON D.nDocTpo =MD.nDocTpo WHERE MD.nMOVNRO=M.nMOVNRO AND MD.nDocTpo ='" & TpoDocSalidaEfectivoBoveda & "' ) ,'')  AS DocSalEfecBov " _
    & " FROM    MOV M " _
    & "         JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO and MC.NMOVIMPORTE >0 " _
    & "         LEFT JOIN MOVHABILITACION MH ON MH.nMOVNRO =M.nMOVNRO " _
    & "         JOIN " & vsServerCom & "AREAS A ON A.CAREACOD = MH.CAREAORIG " _
    & "         LEFT JOIN " & vsServerCom & "AGENCIAS AG ON AG.CAGECOD= MH.CAGEORIG " _
    & "         JOIN " & vsServerCom & "AREAS A1 ON A1.CAREACOD = MH.CAREADEST " _
    & "         LEFT JOIN " & vsServerCom & "AGENCIAS AG1 ON AG1.CAGECOD= MH.CAGEDEST " _
    & "         LEFT JOIN TRANSPORTE T ON T.cTranspCod = MH.cTranspCod " _
    & "         LEFT JOIN " & vsServerPers & "PERSONA P ON P.CPERSCOD = T.CPERSCOD " _
    & "         LEFT JOIN " & vsServerCom & "CONSTANTE C ON C.nConsValor = T.CTPOTRANS AND C.nCONSCOD='" & gCGTipoTransporte & "' " _
    & " WHERE   M.CMOVNRO ='" & psMovNro & "'  AND M.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "') "
    
Set rs = oconect.CargaRecordSet(sql)
Set GetDatosHabilitacion = rs
oconect.CierraConexion
Set oconect = Nothing

End Function
Public Function GetDatosConfHabAgencias(ByVal psMovNro As String) As ADODB.Recordset
Dim sql As String
Dim oconect As DConecta
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
Set oconect = New DConecta
If oconect.AbreConexion = False Then Exit Function

sql = "SELECT   M.CMOVNRO, " _
    & "         MH.CAREAORIG + MH.CAGEORIG + '- ' + ISNULL(AG.CAGEDESCRIPCION,A.CAREADESCRIPCION ) AS ORIGEN, " _
    & "         MH.CAREADEST + MH.CAGEDEST + '- ' + ISNULL(AG1.CAGEDESCRIPCION,A1.CAREADESCRIPCION ) AS DESTINO, " _
    & "         P.CPERSCOD , ISNULL(ISNULL(P.CPERSNOMBRE, T.cDescTrans ),'') AS cNombreTransp, MC.NMOVIMPORTE , ISNULL(MH.nMontoTrans,0) AS nMontoTrans,  ISNULL(C.cConsDescripcion,'') AS cTipoTrans, " _
    & "         ISNULL((SELECT D.CDOCABREV + '-' + MD.CDOCNRO FROM MOVDOC MD JOIN DOCUMENTO D ON D.nDocTpo =MD.nDocTpo WHERE MD.nMOVNRO=M.nMOVNRO AND MD.nDocTpo ='" & TpoDocCompServTranspValores & "' ),'')   AS DocCompTransp, " _
    & "         ISNULL((SELECT D.CDOCABREV + '-' + MD.CDOCNRO FROM MOVDOC MD JOIN DOCUMENTO D ON D.nDocTpo =MD.nDocTpo WHERE MD.nMOVNRO=M.nMOVNRO AND MD.nDocTpo ='" & TpoDocSalidaEfectivoBoveda & "' ) ,'')  AS DocSalEfecBov " _
    & " FROM    MOV M " _
    & "         JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO and MC.NMOVIMPORTE >0 " _
    & "         LEFT JOIN MOVHABILITACION MH ON MH.nMOVNRO =M.nMOVNRO " _
    & "         JOIN " & vsServerCom & "AREAS A ON A.CAREACOD = MH.CAREAORIG " _
    & "         LEFT JOIN " & vsServerCom & "AGENCIAS AG ON AG.CAGECOD= MH.CAGEORIG " _
    & "         JOIN " & vsServerCom & "AREAS A1 ON A1.CAREACOD = MH.CAREADEST " _
    & "         LEFT JOIN " & vsServerCom & "AGENCIAS AG1 ON AG1.CAGECOD= MH.CAGEDEST " _
    & "         LEFT JOIN TRANSPORTE T ON T.cTranspCod = MH.cTranspCod " _
    & "         LEFT JOIN " & vsServerPers & "PERSONA P ON P.CPERSCOD = T.CPERSCOD " _
    & "         LEFT JOIN " & vsServerCom & "CONSTANTE C ON C.nConsValor = T.CTPOTRANS AND C.nCONSCOD='" & gCGTipoTransporte & "' " _
    & "         JOIN MOVREF MR ON MR.NMOVNROREF=M.NMOVNRO JOIN MOV M1 ON M1.NMOVNRO = MR.NMOVNRO  " _
    & " WHERE   M1.CMOVNRO ='" & psMovNro & "'  AND M.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "') "
    
Set rs = oconect.CargaRecordSet(sql)
Set GetDatosConfHabAgencias = rs
oconect.CierraConexion
Set oconect = Nothing

End Function

Public Function GetHabCajaGen(ByVal psCtaCont As String, ByVal psAreaOrig As String, _
                                ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                Optional ByVal psAreaCod As String = "", Optional ByVal psAgeCod As String = "", _
                                Optional ByVal pbMuestraOrigen As Boolean = True) As ADODB.Recordset
Dim sql As String
Dim oconect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltroAge As String

Set rs = New ADODB.Recordset

Set oconect = New DConecta
If oconect.AbreConexion = False Then Exit Function

If psAreaCod <> "" Then
    lsFiltroAge = " AND MH.CAREADEST ='" & psAreaCod & "' AND MH.CAGEDEST='" & psAgeCod & "' "
End If

sql = " SELECT  CONVERT(CHAR(12),CONVERT(DATETIME, SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHAHAB, " _
    & "         " & IIf(pbMuestraOrigen, " ISNULL(AG1.CAGEDESCRIPCION, A1.CAREADESCRIPCION) AS Origen, ", " ISNULL(AG.CAGEDESCRIPCION, A.CAREADESCRIPCION ) AS Destino, ") _
    & "         ISNULL(MH.NMOVIMPORTE,0 )  AS MontoHabilitado , " _
    & "         CONVERT(CHAR(25),C.CCONSDESCRIPCION) as TipoTransp, " _
    & "         CONVERT(CHAR(50), ISNULL(P.CPERSNOMBRE,T.cDescTrans)) as Empresa, MH.nMontoTrans, M.cMovDesc, " _
    & "          M.cMovNro, A.CAREACOD,M.nMovNro , AG.CAGECOD  " _
    & " FROM    MOV M   " _
    & "         JOIN MOVHABILITACION MH ON MH.nMOVNRO=M.nMOVNRO " _
    & "         JOIN " & vsServerCom & "AREAS A1 ON A1.CAREACOD = MH.CAREAORIG " _
    & "         LEFT JOIN " & vsServerCom & "AGENCIAS AG1 ON AG1.CAGECOD= MH.CAGEORIG " _
    & "         JOIN TRANSPORTE T ON T.CTRANSPCOD = MH.CTRANSPCOD " _
    & "         LEFT JOIN " & vsServerPers & "PERSONA P ON P.CPERSCOD = T.CPERSCOD " _
    & "         JOIN " & vsServerCom & "CONSTANTE C ON C.nConsValor = T.cTpoTrans " _
    & "         JOIN " & vsServerCom & "AREAS A ON A.CAREACOD =MH.cAreaOrig " _
    & "         LEFT JOIN " & vsServerCom & "AGENCIAS AG ON AG.CAGECOD = MH.cAgeOrig  " _
    & " WHERE   C.nCONSCOD ='" & gCGTipoTransporte & "' " _
    & "         AND M.nMovFlag NOT IN(" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " & lsFiltroAge _
    & "         AND MH.CAREAORIG ='" & Mid(psAreaOrig, 1, 3) & "' AND MH.CAGEORIG ='" & Mid(psAreaOrig, 4, 2) & "'  " _
    & "         AND NOT EXISTS (    SELECT  M1.CMOVNRO FROM    MOV M1 " _
    & "                                     JOIN MOVREF MR ON M1.nMOVNRO =MR.nMOVNRO " _
    & "                             WHERE MR.nMOVNROREF = M.nMOVNRO AND M1.nMovFlag NOT IN(" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ")) " _
    & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND  '" & Format(pdHasta, "yyyymmdd") & "'"
    
Set rs = oconect.CargaRecordSet(sql)
Set GetHabCajaGen = rs
oconect.CierraConexion
Set oconect = Nothing

End Function
Public Function ExtornaHabEfectivo(ByVal psFormatoFecha As String, _
                                    ByVal pdFecSis As Date, ByVal pdFechaHab As Date, ByVal psMovNroExt As String, _
                                    ByVal pnMovNroHab As Long, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                    ByVal pnImporte As Currency)

Dim oMov As DMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New DMov

On Error GoTo ExtornaHabEfectivoErr
oMov.Inicio psFormatoFecha
oMov.BeginTrans
lbTrans = True
If pdFecSis = pdFechaHab Then
    oMov.InsertaMov psMovNroExt, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagDeExtorno
    lnMovNro = oMov.GetnMovNro(psMovNroExt)
    oMov.InsertaMovCont lnMovNro, pnImporte, 0, "0"
    oMov.ActualizaMov pnMovNroHab, , gMovEstContabMovContable, gMovFlagEliminado
Else
    oMov.ExtornaMovimiento psMovNroExt, pnMovNroHab
End If
oMov.InsertaMovRef pnMovNroHab, lnMovNro
oMov.CommitTrans
lbTrans = False
ExtornaHabEfectivo = 0
Set oMov = Nothing
Exit Function
ExtornaHabEfectivoErr:
    If lbTrans Then
        oMov.RollBackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GrabaConfHabEfectivo(ByVal psMovNroConf As String, _
                                    ByVal psFormatoFecha As String, ByVal rsBillete As ADODB.Recordset, _
                                    ByVal rsMoneda As ADODB.Recordset, _
                                    ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                    ByVal psCtaDebe As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                    ByVal psAreaOrig As String, ByVal psAreaDest As String, _
                                    ByVal psMovNroHab As String) As Integer
            
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim oContF As NContFunciones
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New DMov
On Error GoTo GrabaHabEfectivoErr

Set oContF = New NContFunciones

GrabaConfHabEfectivo = 1
oDMov.Inicio psFormatoFecha
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNroConf, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNroConf)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el haber
lnMovItem = 0: lnMovOrden = 0
'guardamos la cuenta en el debe
If Not rsBillete Is Nothing Then
    If rsBillete.State = adStateOpen Then
        If Not rsBillete.EOF And Not rsBillete.BOF Then
            Do While Not rsBillete.EOF
                If rsBillete!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaDebe, psAreaDest, False)
                    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe + lsSubCta, rsBillete!Monto
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oDMov.InsertaMovObjEfectivo lnMovNro, Format(lnMovItem, "#0"), Format(lnMovOrden, "#0"), rsBillete!cEfectivoCod
                End If
                rsBillete.MoveNext
            Loop
        End If
        rsBillete.Close: Set rsBillete = Nothing
    End If
End If
If Not rsMoneda Is Nothing Then
    If rsMoneda.State = adStateOpen Then
        If Not rsMoneda.EOF And Not rsMoneda.BOF Then
            Do While Not rsMoneda.EOF
                If rsMoneda!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaDebe, psAreaDest, False)
                    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe + lsSubCta, rsMoneda!Monto
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oDMov.InsertaMovObjEfectivo lnMovNro, Format(lnMovItem, "#0"), Format(lnMovOrden, "#0"), rsMoneda!cEfectivoCod
                End If
                rsMoneda.MoveNext
            Loop
        End If
        rsMoneda.Close: Set rsMoneda = Nothing
    End If
End If
'insertamos la cuenta del haber
lnMovItem = lnMovItem + 1: lnMovOrden = 0
lnMovOrden = lnMovOrden + 1
lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaHaber, psAreaOrig, False)
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber + lsSubCta, pnImporte * -1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
oDMov.InsertaMovObjAgenciaArea lnMovNro, Format(lnMovItem, "#0"), Format(lnMovOrden, "#0"), Mid(psAreaOrig, 4, 2), Mid(psAreaOrig, 1, 3)

oDMov.InsertaMovRef lnMovNro, psMovNroHab

oDMov.CommitTrans
lbTrans = False
GrabaConfHabEfectivo = 0
Set oDMov = Nothing
Set oContF = Nothing
Exit Function
GrabaHabEfectivoErr:
    If lbTrans Then
        oDMov.RollBackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaHabEfectivo", Err.Description & " Error Graba Habilitación de Efectivo"
End Function
Public Function GetDatosConfHabilitacion(ByVal psCtaCont As String, ByVal psAreaOrig As String, _
                                        ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                        Optional ByVal psAreaCod As String = "", Optional ByVal psAgeCod As String = "") As ADODB.Recordset
Dim sql As String
Dim oconect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltroAge As String
Set rs = New ADODB.Recordset
Set oconect = New DConecta
If oconect.AbreConexion = False Then Exit Function

If psAreaCod <> "" Then
    lsFiltroAge = " AND MOAA.CAREACOD ='" & psAreaCod & "' AND MOAA.CAGECOD='" & psAgeCod & "' "
End If


sql = " SELECT  CONVERT(CHAR(12),CONVERT(DATETIME, SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHAHAB," _
    & "         ISNULL(AG1.CAGEDESCRIPCION, A1.CAREADESCRIPCION ) AS Destino,  " _
    & "         ISNULL(ME.NMOVMEIMPORTE , MC.NMOVIMPORTE )  AS MontoHabilitado , " _
    & "         CONVERT(CHAR(25),C.CCONSDESCRIPCION) as TipoTransp, " _
    & "         CONVERT(CHAR(50), ISNULL(P.CPERSNOMBRE,T.cDescTrans)) as Empresa, " _
    & "         MH.nMontoTrans, M.cMovDesc, CONF.CMOVNRO AS CMOVCONFHAB , MOAA.CAREACOD, " _
    & "         CONF.nMovNro, MOAA.CAGECOD , M.cMovNro " _
    & " FROM    MOV M " _
    & "         JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO " _
    & "         LEFT JOIN MOVME ME ON ME.nMOVNRO =MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "         JOIN MOVHABILITACION MH ON MH.nMOVNRO=M.nMOVNRO " _
    & "         JOIN " & vsServerCom & "AREAS A1 ON A1.CAREACOD = MH.CAREAORIG " _
    & "         LEFT JOIN " & vsServerCom & "AGENCIAS AG1 ON AG1.CAGECOD= MH.CAGEORIG " _
    & "         JOIN TRANSPORTE T ON T.CTRANSPCOD = MH.CTRANSPCOD " _
    & "         LEFT JOIN " & vsServerPers & "PERSONA P ON P.CPERSCOD = T.CPERSCOD " _
    & "         JOIN " & vsServerCom & "CONSTANTE C ON C.nConsValor = T.cTpoTrans " _
    & "         JOIN MOVOBJ MO ON MO.nMOVNRO =MC.nMOVNRO AND MO.nMOVITEM = MC.nMOVITEM " _
    & "         JOIN MOVOBJAREAAGENCIA MOAA ON MOAA.nMOVNRO = MO.nMOVNRO AND MOAA.nMOVITEM = MO.nMOVITEM AND MOAA.nMovObjOrden = MO.nMovObjOrden " _
    & "         JOIN " & vsServerCom & "AREAS A ON MOAA.CAREACOD= A.CAREACOD " _
    & "         LEFT JOIN " & vsServerCom & "AGENCIAS AG ON AG.CAGECOD = MOAA.CAGECOD "
sql = sql + "   JOIN (  SELECT  M1.CMOVNRO, MR.nMOVNROREF , M1.nMOVNRO " _
    & "                 FROM    MOV M1 " _
    & "                         JOIN MOVREF MR ON M1.nMOVNRO =MR.nMOVNRO " _
    & "                 WHERE   M1.nMovFlag NOT IN(" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ")) AS CONF " _
    & "         ON CONF.nMOVNROREF = M.nMOVNRO " _
    & " WHERE   MC.CCTACONTCOD LIKE  '" & psCtaCont & "%' AND C.nCONSCOD ='" & gCGTipoTransporte & "'" _
    & "         AND M.nMovFlag NOT  IN(" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " & lsFiltroAge _
    & "         AND MH.CAREAORIG ='" & Mid(psAreaOrig, 1, 3) & "' AND MH.CAGEORIG ='" & Mid(psAreaOrig, 4, 2) & "' " _
    & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND  '" & Format(pdHasta, "yyyymmdd") & "'"
    

    
Set rs = oconect.CargaRecordSet(sql)
Set GetDatosConfHabilitacion = rs
oconect.CierraConexion
Set oconect = Nothing

End Function
Public Function GrabaCompraVentaMEIF(ByVal psFormatoFecha As String, ByVal psMovNro As String, _
                                ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                ByVal psCtaDebe As String, ByVal psCtaHaber As String, _
                                ByVal psCtaIfOrig As String, ByVal psCtaIfDest As String, _
                                ByVal psTpoDocOrig As TpoDoc, ByVal psNroDocOrig As String, _
                                ByVal psTpoDocDest As TpoDoc, ByVal psNroDocDest As String, _
                                ByVal pnImporte As Currency, ByVal pnTCBanco As Currency, ByVal pdFecSis As Date)

Dim oMov As DMov
Dim lnItem As Integer
Dim lnOrden As Integer
Dim lbTrans As Boolean
Dim lnMovNro As Long
On Error GoTo GrabaCompraVentaMEIFErr

Set oMov = New DMov
oMov.Inicio psFormatoFecha
oMov.BeginTrans
lbTrans = True
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc
lnMovNro = oMov.GetnMovNro(psMovNro)

oMov.InsertaMovCont lnMovNro, pnImporte, 0, "0"
lnItem = 0
lnItem = lnItem + 1: lnOrden = 0
oMov.InsertaMovCta lnMovNro, lnItem, psCtaDebe, pnImporte
lnOrden = lnOrden + 1
oMov.InsertaMovObj lnMovNro, lnItem, lnOrden, ObjEntidadesFinancieras
oMov.InsertaMovObjIF lnMovNro, lnItem, lnOrden, Mid(psCtaIfOrig, 4, 13), Mid(psCtaIfOrig, 1, 2), Mid(psCtaIfOrig, 18, 10)

'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
oMov.InsertaMovCtaIF lnMovNro, Mid(psCtaIfOrig, 4, 13), Mid(psCtaIfOrig, 1, 2), Mid(psCtaIfOrig, 18, 10), CGCtaIFConCapital, pnImporte
oMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                        Mid(psCtaIfOrig, 4, 13), Mid(psCtaIfOrig, 1, 2), Mid(psCtaIfOrig, 18, 10)
                        
lnItem = lnItem + 1: lnOrden = 0
oMov.InsertaMovCta lnMovNro, lnItem, psCtaHaber, pnImporte * -1
lnOrden = lnOrden + 1
oMov.InsertaMovObj lnMovNro, lnItem, lnOrden, ObjEntidadesFinancieras
oMov.InsertaMovObjIF lnMovNro, lnItem, lnOrden, Mid(psCtaIfDest, 4, 13), Mid(psCtaIfDest, 1, 2), Mid(psCtaIfDest, 18, 10)

'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
oMov.InsertaMovCtaIF lnMovNro, Mid(psCtaIfDest, 4, 13), Mid(psCtaIfDest, 1, 2), Mid(psCtaIfDest, 18, 10), CGCtaIFConCapital, pnImporte * -1
oMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                        Mid(psCtaIfDest, 4, 13), Mid(psCtaIfDest, 1, 2), Mid(psCtaIfDest, 18, 10)

Select Case psOpeCod
    Case gOpeMECompraAInst, gOpeMEVentaAInst
        oMov.InsertaMovTpoCambio lnMovNro, pnTCBanco
        'omov.GeneraMovME lsmovnro
End Select
If psNroDocOrig <> "" Then
    oMov.InsertaMovDoc lnMovNro, psTpoDocOrig, psNroDocOrig, pdFecSis
End If
If psNroDocDest <> "" Then
    oMov.InsertaMovDoc lnMovNro, psTpoDocDest, psNroDocDest, pdFecSis
End If

oMov.CommitTrans
lbTrans = False
Set oMov = Nothing
Exit Function
GrabaCompraVentaMEIFErr:
        If lbTrans Then
            oMov.RollBackTrans
        End If
        RaiseError Err.Number, Err.Description
End Function
Public Function GrabaCompraVentaEfectivo(ByVal psFormatoFecha As String, _
                                    ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                    ByVal rsBillIngreso As ADODB.Recordset, ByVal rsMonIng As ADODB.Recordset, _
                                    ByVal rsBillSale As ADODB.Recordset, ByVal rsMonSale As ADODB.Recordset, _
                                    ByVal psCtaDebe As String, ByVal psCtaHaber As String, _
                                    ByVal pnTCBanco As Currency) As Integer
            
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim oContF As NContFunciones
Dim lsMovNro As String
Dim lnMovNro As Long
Dim lsSubCta As String
Dim lbTrans As Boolean

Set oDMov = New DMov
On Error GoTo GrabaCompraVentaEfectivoErr

Set oContF = New NContFunciones

GrabaCompraVentaEfectivo = 1
oDMov.Inicio psFormatoFecha
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, 0, "0", "0"
'guardamos la cuenta en el haber
lnMovItem = 0: lnMovOrden = 0
'guardamos la cuenta en el debe
If Not rsBillIngreso Is Nothing Then
    If rsBillIngreso.State = adStateOpen Then
        If Not rsBillIngreso.EOF And Not rsBillIngreso.BOF Then
            Do While Not rsBillIngreso.EOF
                If rsBillIngreso!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe, rsBillIngreso!Monto
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, rsBillIngreso!cEfectivoCod
                End If
                rsBillIngreso.MoveNext
            Loop
        End If
        rsBillIngreso.Close: Set rsBillIngreso = Nothing
    End If
End If
If Not rsMonIng Is Nothing Then
    If rsMonIng.State = adStateOpen Then
        If Not rsMonIng.EOF And Not rsMonIng.BOF Then
            Do While Not rsMonIng.EOF
                If rsMonIng!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, rsMonIng!Monto
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oDMov.InsertaMovObjEfectivo lnMovNro, Format(lnMovItem, "#0"), Format(lnMovOrden, "#0"), rsMonIng!cEfectivoCod
                End If
                rsMonIng.MoveNext
            Loop
        End If
        rsMonIng.Close: Set rsMonIng = Nothing
    End If
End If
'gaurdamos las cuentas del haber
If Not rsBillSale Is Nothing Then
    If rsBillSale.State = adStateOpen Then
        If Not rsBillSale.EOF And Not rsBillSale.BOF Then
            Do While Not rsBillSale.EOF
                If rsBillSale!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, rsBillSale!Monto * -1
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, rsBillSale!cEfectivoCod
                End If
                rsBillSale.MoveNext
            Loop
        End If
        rsBillSale.Close: Set rsBillSale = Nothing
    End If
End If
If Not rsMonIng Is Nothing Then
    If rsMonIng.State = adStateOpen Then
        If Not rsMonIng.EOF And Not rsMonIng.BOF Then
            Do While Not rsMonIng.EOF
                If rsMonIng!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber * -1, rsMonIng!Monto
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, rsMonIng!cEfectivoCod
                End If
                rsMonIng.MoveNext
            Loop
        End If
        rsMonIng.Close: Set rsMonIng = Nothing
    End If
End If
oDMov.InsertaMovTpoCambio lnMovNro, pnTCBanco

GrabaCompraVentaEfectivo = 0
oDMov.CommitTrans
lbTrans = False
Set oDMov = Nothing
Set oContF = Nothing
Exit Function
GrabaCompraVentaEfectivoErr:
    If lbTrans Then
        oDMov.RollBackTrans
    End If
    Err.Raise vbObjectError + 100, "GrabaCompraVentaEfectivo", Err.Description & " Error Graba Compra Venta Efectivo"
End Function
Public Function GetMovCompraVentaME(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset
Dim sql As String
Dim oconect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltroAge As String
Set rs = New ADODB.Recordset
Set oconect = New DConecta
If oconect.AbreConexion = False Then Exit Function


sql = " SELECT  CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) as Fecha ," _
    & "         D.CPERSNOMBRE, D.CTAIFDEST, H.cPersNombre , H.CTAIFORIG, D.IMPORTE, M.nMovNro, M.cMovDesc " _
    & " FROM    MOV M " _
    & "         JOIN (  SELECT  M.CMOVNRO , M.nMOVNRO  , P.cPersNombre , C1.cCtaIFDesc ,  C.cCtaIFDesc AS CTAIFORIG , " _
    & "                         ISNULL(Me.NMOVMEIMPORTE, MC.NMOVIMPORTE) As IMPORTE " _
    & "                 FROM    MOV M  JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO  " _
    & "                         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "                         JOIN MOVOBJIF MO ON MO.nMOVNRO = MC.nMOVNRO AND MO.nMOVITEM = MC.nMOVITEM " _
    & "                         JOIN CTAIF C ON C.CPERSCOD = MO.CPERSCOD AND MO.CCTAIFCOD = C.CCTAIFCOD  AND C.CIFTPO = MO.CIFTPO   " _
    & "                         JOIN CTAIF C1 ON C1.CPERSCOD = MO.CPERSCOD AND C1.CIFTPO = MO.CIFTPO AND MO.CCTAIFCOD LIKE C1.cCtaIFCod+'%' AND LEN(C1.cCtaIFCod) < LEN (MO.CCTAIFCOD) " _
    & "                         JOIN " & vsServerPers & "PERSONA P ON P.CPERSCOD = C.CPERSCOD " _
    & "                 WHERE   MC.NMOVIMPORTE <0  AND M.COPECOD ='" & psOpeCod & "'  ) AS H " _
    & "         ON H.nMOVNRO = M.nMOVNRO " _
    & "         JOIN ( SELECT   M.CMOVNRO, M.nMOVNRO ,  P.cPersNombre,  C1.cCtaIFDesc  ,  C.cCtaIFDesc  AS CTAIFDEST, " _
    & "                         ISNULL(Me.NMOVMEIMPORTE, MC.NMOVIMPORTE) As IMPORTE " _
    & "                 FROM    MOV M  JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO  " _
    & "                         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "                         JOIN MOVOBJIF MO ON MO.nMOVNRO = MC.nMOVNRO AND MO.nMOVITEM = MC.nMOVITEM " _
    & "                         JOIN CTAIF C ON C.CPERSCOD = MO.CPERSCOD AND MO.CCTAIFCOD = C.CCTAIFCOD AND C.CIFTPO = MO.CIFTPO " _
    & "                         JOIN CTAIF C1 ON C1.CPERSCOD = MO.CPERSCOD AND C1.CIFTPO = MO.CIFTPO AND MO.CCTAIFCOD LIKE C1.cCtaIFCod+'%' AND LEN(C1.cCtaIFCod) < LEN (MO.CCTAIFCOD) " _
    & "                         JOIN " & vsServerPers & "PERSONA P ON P.CPERSCOD = C.CPERSCOD " _
    & "                 WHERE   MC.NMOVIMPORTE >0 AND M.COPECOD ='" & psOpeCod & "' ) AS D " _
    & "         ON D.nMOVNRO = M.nMOVNRO " _
    & " WHERE   M.COPECOD ='" & psOpeCod & "' AND M.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "')" _
    & "         AND SUBSTRING(M.cMovnro,1,8 ) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "'  ORDER BY M.CMOVNRO "


Set rs = oconect.CargaRecordSet(sql)
Set GetMovCompraVentaME = rs
oconect.CierraConexion
Set oconect = Nothing

End Function
Public Function GrabaExtornoMov(ByVal psFormatoFecha As String, _
                                  ByVal pdFecSis As Date, ByVal pdFechaMov As Date, ByVal psMovNroExt As String, _
                                  ByVal pnMovNroAnt As Long, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                  ByVal pnImporte As Currency)

On Error GoTo ErrorGrabaExtornoMov
Dim oMov As DMov
Dim lnMovNro As Long
Dim lbTrans As Boolean
Set oMov = New DMov

oMov.BeginTrans
lbTrans = True
oMov.Inicio psFormatoFecha
If pdFecSis = Format(pdFechaMov, gsFormatoFechaView) Then
    oMov.InsertaMov psMovNroExt, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagDeExtorno
    lnMovNro = oMov.GetnMovNro(psMovNroExt)
    oMov.InsertaMovCont lnMovNro, pnImporte, 0, "0"
    oMov.ActualizaMov pnMovNroAnt, , gMovEstContabNoContable, gMovFlagEliminado
Else
    oMov.ExtornaMovimiento psMovNroExt, pnMovNroAnt
    lnMovNro = oMov.GetnMovNro(psMovNroExt)
End If
oMov.InsertaMovRef pnMovNroAnt, lnMovNro
oMov.CommitTrans
lbTrans = False
Set oMov = Nothing
Exit Function
ErrorGrabaExtornoMov:
    If lbTrans Then oMov.RollBackTrans
    Err.Raise Err.Number, "GrabaExtornoMov", Err.Description
End Function
Public Function GetMovCompraVentaEfectivo(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset
Dim sql As String
Dim oconect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltroAge As String
Set rs = New ADODB.Recordset
Set oconect = New DConecta
If oconect.AbreConexion = False Then Exit Function


sql = " SELECT  CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) as Fecha ," _
    & "         M.CMOVDESC, D.MEIMPORTE AS DOLARES, S.MNIMPORTE , MTC.nMovTpoCambio, S.nMOVNRO " _
    & " FROM    MOV M " _
    & "         JOIN (  SELECT  M.CMOVNRO, M.nMOVNRO , SUM(ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE))) AS MEIMPORTE " _
    & "                 FROM    MOV M JOIN MOVCTA MC  ON MC.nMOVNRO =M.nMOVNRO " _
    & "                         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "                 WHERE   MC.CCTACONTCOD LIKE '__" & gMonedaExtranjera & "%'  AND M.COPECOD ='" & psOpeCod & "'  " _
    & "                 GROUP  BY  M.cMOVNRO, M.nMOVNRO   ) AS D " _
    & "         ON D.nMOVNRO = M.nMOVNRO " _
    & "         JOIN (  SELECT  M.CMOVNRO, M.nMOVNRO, SUM(ISNULL(ABS(MC.NMOVIMPORTE),0)) AS MNIMPORTE " _
    & "                 FROM    MOV M JOIN MOVCTA MC  ON MC.nMOVNRO =M.nMOVNRO " _
    & "                         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "                 WHERE   MC.CCTACONTCOD LIKE '__" & gMonedaNacional & "%' AND M.COPECOD ='" & psOpeCod & "'  " _
    & "                 GROUP  BY  M.CMOVNRO, M.nMOVNRO   ) AS S " _
    & "         ON S.nMOVNRO = M.nMOVNRO " _
    & "         JOIN MOVTPOCAMBIO  MTC ON MTC.nMOVNRO = M.nMOVNRO " _
    & " WHERE   M.COPECOD ='" & psOpeCod & "' AND M.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "')" _
    & "         AND SUBSTRING(S.cMovnro,1,8 ) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "'  ORDER BY M.CMOVNRO "

Set rs = oconect.CargaRecordSet(sql)
Set GetMovCompraVentaEfectivo = rs
oconect.CierraConexion
Set oconect = Nothing

End Function
Public Function GrabaMovimientosCtasIF(ByVal psFormatoFecha As String, ByVal psMovNro As String, _
                                        ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                        ByVal psCtaDebe As String, ByVal psCtaHaber As String, _
                                        ByVal psCtaIfOrig As String, ByVal psCtaIfDest As String, _
                                        ByVal psTpoDocOrig As TpoDoc, ByVal psNroDocOrig As String, _
                                        ByVal psTpoDocDest As TpoDoc, ByVal psNroDocDest As String, _
                                        ByVal pnImporte As Currency, ByVal pnTCBanco As Currency, ByVal pdFecSis As Date)

Dim oMov As DMov
Dim lnItem As Integer
Dim lnOrden As Integer
Dim lbTrans As Boolean
Dim lnMovNro As Long
On Error GoTo GrabaMovimientosCtasIFErr

Set oMov = New DMov

GrabaMovimientosCtasIF = 1

oMov.BeginTrans
oMov.Inicio psFormatoFecha
lbTrans = True
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc
lnMovNro = oMov.GetnMovNro(psMovNro)

oMov.InsertaMovCont lnMovNro, pnImporte, 0, "0"
lnItem = 0
lnItem = lnItem + 1: lnOrden = 0
oMov.InsertaMovCta lnMovNro, lnItem, psCtaDebe, pnImporte
lnOrden = lnOrden + 1
oMov.InsertaMovObj lnMovNro, lnItem, lnOrden, ObjEntidadesFinancieras
oMov.InsertaMovObjIF lnMovNro, lnItem, lnOrden, Mid(psCtaIfOrig, 4, 13), Mid(psCtaIfOrig, 1, 2), Mid(psCtaIfOrig, 18, 10)

'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
oMov.InsertaMovCtaIF lnMovNro, Mid(psCtaIfOrig, 4, 13), Mid(psCtaIfOrig, 1, 2), Mid(psCtaIfOrig, 18, 10), CGCtaIFConCapital, pnImporte
oMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                        Mid(psCtaIfOrig, 4, 13), Mid(psCtaIfOrig, 1, 2), Mid(psCtaIfOrig, 18, 10)



lnItem = lnItem + 1: lnOrden = 0
oMov.InsertaMovCta lnMovNro, lnItem, psCtaHaber, pnImporte * -1
lnOrden = lnOrden + 1
oMov.InsertaMovObj lnMovNro, lnItem, lnOrden, ObjEntidadesFinancieras
oMov.InsertaMovObjIF lnMovNro, lnItem, lnOrden, Mid(psCtaIfDest, 4, 13), Mid(psCtaIfDest, 1, 2), Mid(psCtaIfDest, 18, 10)

'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
oMov.InsertaMovCtaIF lnMovNro, Mid(psCtaIfDest, 4, 13), Mid(psCtaIfDest, 1, 2), Mid(psCtaIfDest, 18, 10), CGCtaIFConCapital, pnImporte * -1
oMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                        Mid(psCtaIfDest, 4, 13), Mid(psCtaIfDest, 1, 2), Mid(psCtaIfDest, 18, 10)

If pnTCBanco <> 0 Then
    oMov.InsertaMovTpoCambio lnMovNro, pnTCBanco
End If
If psNroDocOrig <> "" Then
    oMov.InsertaMovDoc lnMovNro, psTpoDocOrig, psNroDocOrig, pdFecSis
End If
If psNroDocDest <> "" Then
    oMov.InsertaMovDoc lnMovNro, psTpoDocDest, psNroDocDest, pdFecSis
End If

oMov.CommitTrans
lbTrans = False
Set oMov = Nothing
GrabaMovimientosCtasIF = 0
Exit Function
GrabaMovimientosCtasIFErr:
        If lbTrans Then
            oMov.RollBackTrans
            lbTrans = False
        End If
        RaiseError Err.Number, Err.Description
End Function
Public Function GrabaMovEfectivo(ByVal psFormatoFecha As String, _
                                ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                ByVal rsBill As ADODB.Recordset, ByVal rsMon As ADODB.Recordset, _
                                ByVal psCtaCont As String, _
                                ByVal psCuentaEfect As String, ByVal pnImporte As Currency, _
                                ByVal psTipoObj As TpoObjetos, ByVal psObjeto As String, _
                                ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pdFechaDoc As Date, _
                                Optional ByVal pbIngreso As Boolean = False, Optional pnMovRef As Long = -1) As Integer
            
Dim oDMov As DMov
Set oDMov = New DMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lnMovItem As Integer, lnMovOrden As Integer
On Error GoTo GrabaMovEfectivoErr

GrabaMovEfectivo = 1
oDMov.BeginTrans
lbTrans = True
oDMov.Inicio psFormatoFecha
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
lnMovItem = lnMovItem + 1
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaCont, pnImporte * IIf(pbIngreso = False, 1, -1)
If psTipoObj <> -1 Then
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, psTipoObj
Select Case psTipoObj
    Case ObjCMACAgenciaArea
        oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(psObjeto, 4, 2), Mid(psObjeto, 1, 3)
    Case ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psObjeto, 4, 13), Mid(psObjeto, 1, 2), Mid(psObjeto, 18, 10)
        
        'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
        oDMov.InsertaMovCtaIF lnMovNro, Mid(psObjeto, 4, 13), Mid(psObjeto, 1, 2), Mid(psObjeto, 18, 10), CGCtaIFConCapital, pnImporte * IIf(pbIngreso = False, 1, -1)
        oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                                Mid(psObjeto, 4, 13), Mid(psObjeto, 1, 2), Mid(psObjeto, 18, 10)
    Case ObjDescomEfectivo
        oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, psObjeto
    Case ObjPersona
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, psObjeto
    Case Else
End Select
End If
If Not rsBill Is Nothing Then
    If rsBill.State = adStateOpen Then
        If Not rsBill.EOF And Not rsBill.BOF Then
            Do While Not rsBill.EOF
                If rsBill!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCuentaEfect, rsBill!Monto * IIf(pbIngreso = False, -1, 1)
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, rsBill!cEfectivoCod
                End If
                rsBill.MoveNext
            Loop
        End If
        rsBill.Close: Set rsBill = Nothing
    End If
End If
If Not rsMon Is Nothing Then
    If rsMon.State = adStateOpen Then
        If Not rsMon.EOF And Not rsMon.BOF Then
            Do While Not rsMon.EOF
                If rsMon!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCuentaEfect, rsMon!Monto * IIf(pbIngreso = False, -1, 1)
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, rsMon!cEfectivoCod
                End If
                rsMon.MoveNext
            Loop
        End If
        rsMon.Close: Set rsMon = Nothing
    End If
End If
If psNroDoc <> "" Then
    oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, pdFechaDoc
End If
If pnMovRef <> -1 Then
    oDMov.ActualizaMov pnMovRef, , gMovEstContabMovContable
    oDMov.ExtornaMovCtaIF pnMovRef
    oDMov.InsertaMovRef lnMovNro, pnMovRef
End If

oDMov.CommitTrans
lbTrans = False

GrabaMovEfectivo = 0
Set oDMov = Nothing

Exit Function
GrabaMovEfectivoErr:
    If lbTrans Then
        oDMov.RollBackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaMovEfectivo", Err.Description & " Error Registro de Movimiento de Efectivo de Caja"
End Function
Public Function GrabaMovGeneral(ByVal psFormatoFecha As String, _
                                ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                ByVal psCtaDebe As String, _
                                ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                ByVal pnTipoObjDeb As TpoObjetos, ByVal psObjetoDeb As String, _
                                ByVal pnTipoObjHab As TpoObjetos, ByVal psObjetoHab As String, _
                                ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pdFechaDoc As Date, _
                                ByVal psNroVoucher As String) As Integer
            
Dim oDMov As DMov
Set oDMov = New DMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lnMovItem As Integer, lnMovOrden As Integer
Dim lsFiltro As String
Dim oCont As NContFunciones
On Error GoTo GrabaMovGeneralErr

GrabaMovGeneral = 1
Set oCont = New NContFunciones
oDMov.BeginTrans
lbTrans = True
oDMov.Inicio psFormatoFecha
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
lnMovItem = lnMovItem + 1
lsFiltro = oCont.GetFiltroObjetos(pnTipoObjDeb, psCtaDebe, psObjetoDeb, False)
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe + lsFiltro, pnImporte
If pnTipoObjDeb <> -1 Then
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, pnTipoObjDeb
Select Case pnTipoObjDeb
    Case ObjCMACAgenciaArea
        oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(psObjetoDeb, 4, 2), Mid(psObjetoDeb, 1, 3)
    Case ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psObjetoDeb, 4, 13), Mid(psObjetoDeb, 1, 2), Mid(psObjetoDeb, 18, 10)
        'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
        oDMov.InsertaMovCtaIF lnMovNro, Mid(psObjetoDeb, 4, 13), Mid(psObjetoDeb, 1, 2), Mid(psObjetoDeb, 18, 10), CGCtaIFConCapital, pnImporte
        oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                                Mid(psObjetoDeb, 4, 13), Mid(psObjetoDeb, 1, 2), Mid(psObjetoDeb, 18, 10)

    Case ObjDescomEfectivo
        oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, psObjetoDeb
    Case ObjPersona
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, psObjetoDeb
    Case Else
End Select
End If
lnMovItem = lnMovItem + 1: lnMovOrden = 0
lsFiltro = oCont.GetFiltroObjetos(pnTipoObjHab, psCtaHaber, psObjetoHab, False)
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber + lsFiltro, pnImporte * -1
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, pnTipoObjHab
If pnTipoObjHab <> -1 Then
Select Case pnTipoObjHab
    Case ObjCMACAgenciaArea
        oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(psObjetoHab, 4, 2), Mid(psObjetoHab, 1, 3)
    Case ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psObjetoHab, 4, 13), Mid(psObjetoHab, 1, 2), Mid(psObjetoHab, 18, 10)
        'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
        oDMov.InsertaMovCtaIF lnMovNro, Mid(psObjetoHab, 4, 13), Mid(psObjetoHab, 1, 2), Mid(psObjetoHab, 18, 10), CGCtaIFConCapital, pnImporte * -1
        
        oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                                Mid(psObjetoHab, 4, 13), Mid(psObjetoHab, 1, 2), Mid(psObjetoHab, 18, 10)

    Case ObjDescomEfectivo
        oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, psObjetoHab
    Case ObjPersona
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, psObjetoHab
    Case Else
End Select
End If
If psNroDoc <> "" Then
    oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, pdFechaDoc
End If
If psNroVoucher <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psNroVoucher, pdFechaDoc
End If

oDMov.CommitTrans
lbTrans = False

GrabaMovGeneral = 0
Set oDMov = Nothing

Exit Function
GrabaMovGeneralErr:
    If lbTrans Then
        oDMov.RollBackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaMovEfectivo", Err.Description & " Error Registro de Movimiento de Efectivo de Caja"
End Function
Public Function GetRetirosEfectivoCaja(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset
Dim sql As String
Dim oconect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltroAge As String
Set rs = New ADODB.Recordset
Set oconect = New DConecta
If oconect.AbreConexion = False Then Exit Function



sql = " SELECT  DOCMOV.cDocAbrev, DOCMOV.cDocNro, VOUC.cDocNro AS VOUCHER, P.cPersNombre, " _
    & "         RTRIM(C1.cCtaIFDesc) + ' ' + C.cCtaIFDesc AS CUENTA, " _
    & "         ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE))  AS IMPORTE, M.NMOVNRO, M.CMOVDESC " _
    & " FROM    MOV M " _
    & "         JOIN MOVCTA MC ON MC.NMOVNRO =M.NMOVNRO " _
    & "         LEFT JOIN MOVME ME ON ME.NMOVNRO =MC.NMOVNRO AND ME.NMOVITEM=MC.NMOVITEM " _
    & "         JOIN (  SELECT  MD.NMOVNRO, D.cDocAbrev, MD.cDocNro " _
    & "                 FROM    MOVDOC MD JOIN DOCUMENTO D ON D.NDOCTPO=MD.NDOCTPO " _
    & "                 WHERE   MD.nDocTpo<>" & TpoDocVoucherEgreso & " ) AS DOCMOV " _
    & "         ON DOCMOV.NMOVNRO=M.NMOVNRO " _
    & "         LEFT JOIN (  SELECT  MD.NMOVNRO, MD.cDocNro " _
    & "                 FROM    MOVDOC MD " _
    & "                 WHERE MD.nDocTpo=" & TpoDocVoucherEgreso & " ) AS VOUC " _
    & "         ON VOUC.NMOVNRO=M.NMOVNRO " _
    & "         JOIN MOVOBJIF MIF ON MIF.NMOVNRO=MC.NMOVNRO AND MC.nMovItem =MIF.nMovItem " _
    & "         JOIN PERSONA P ON P.CPERSCOD = MIF.CPERSCOD " _
    & "         JOIN CTAIF C1 ON SUBSTRING(MIF.cCtaIFCod,1,LEN(C1.cCtaIFCod))=C1.cCtaIFCod AND LEN(C1.cCtaIFCod)=5 " _
    & "                             AND MIF.CPERSCOD=C1.CPERSCOD " _
    & "         JOIN CTAIF C ON MIF.cCtaIFCod=C.cCtaIFCod AND MIF.CPERSCOD=C.CPERSCOD " _
    & " WHERE   M.COPECOD ='" & psOpeCod & "' AND M.nMovFlag = '" & gMovFlagVigente & "' " _
    & "         AND SUBSTRING(M.cMovnro,1,8 ) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "'  " _
    & "         AND NOT EXISTS (SELECT  M1.CMOVNRO FROM  MOV M1 " _
    & "                                 JOIN MOVREF MR ON M1.nMOVNRO =MR.nMOVNRO " _
    & "                         WHERE MR.nMOVNROREF = M.nMOVNRO AND M1.nMovFlag = " & gMovFlagVigente & ") " _
    & " ORDER BY M.CMOVNRO "

Set rs = oconect.CargaRecordSet(sql)
Set GetRetirosEfectivoCaja = rs
oconect.CierraConexion
Set oconect = Nothing

End Function
Public Function GrabaDepositoCheques(ByVal psFormatoFecha As String, _
                                    ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                    ByVal psCtaDebe As String, ByVal psObjetoIF As String, _
                                    ByVal psCtaHaber As String, ByVal pnTotalImporte As Currency, _
                                    ByVal rs As ADODB.Recordset) As Integer
            
Dim oDMov As DMov
Set oDMov = New DMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lnMovItem As Integer, lnMovOrden As Integer
Dim lsFiltro As String
Dim oCont As NContFunciones
On Error GoTo GrabaDepositoChequesErr

GrabaDepositoCheques = 1
Set oCont = New NContFunciones
oDMov.BeginTrans
lbTrans = True
oDMov.Inicio psFormatoFecha
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
lnMovItem = lnMovItem + 1

oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, pnTotalImporte
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psObjetoIF, 4, 13), Mid(psObjetoIF, 1, 2), Mid(psObjetoIF, 18, 10)

'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
oDMov.InsertaMovCtaIF lnMovNro, Mid(psObjetoIF, 4, 13), Mid(psObjetoIF, 1, 2), Mid(psObjetoIF, 18, 10), CGCtaIFConCapital, pnTotalImporte
oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                            Mid(psObjetoIF, 4, 13), Mid(psObjetoIF, 1, 2), Mid(psObjetoIF, 18, 10)


lnMovItem = lnMovItem + 1: lnMovOrden = 0
If Not rs Is Nothing Then
    If rs.State = adStateOpen Then
        If Not rs.EOF And Not rs.BOF Then
            Do While Not rs.EOF
                If Val(rs!Ok) = 1 Then
                    lsFiltro = ""
                    lsFiltro = oCont.GetFiltroObjetos(ObjProductosCMACT, psCtaHaber, rs!cObjetoCod, False)
                    lsFiltro = lsFiltro & oCont.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaHaber, rs!AreaCod & rs!Agecod, False)
                    lnMovItem = lnMovItem + 1
                    lnMovOrden = 0
                    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber + lsFiltro, rs!Importe * -1
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rs!cObjetoCod
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
                    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, rs!Agecod, rs!AreaCod
                    
                    oDMov.ActualizaCheque TpoDocCheque, rs![N° Cheque], rs!cPersCod, rs!cIFTpo, , , , , , gCGEstadosChqDepositado
                    
                    oDMov.InsertaMovRef lnMovNro, rs!nMovNro
                End If
                rs.MoveNext
            Loop
        End If
    End If
End If
'If psNroDoc <> "" Then
'    oDMov.InsertaMovDoc lnMovNro, pnTpodoc, psNroDoc, pdFechaDoc
'End If

oDMov.CommitTrans
lbTrans = False

GrabaDepositoCheques = 0
Set oDMov = Nothing

Exit Function
GrabaDepositoChequesErr:
    If lbTrans Then
        oDMov.RollBackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaDepositoCheques", Err.Description & " Error Registro de Movimiento de Efectivo de Caja"
End Function
Public Function GetMotivosDepRet(ByVal pnTipoMot As MotivosDepRet) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oconect As DConecta

Set rs = New ADODB.Recordset
Set oconect = New DConecta

sql = " SELECT cMotDesc, cMotCod  " _
    & " FROM MOTIVOSDEPRET WHERE cMotCod LIKE '" & pnTipoMot & "%' AND LEN(CMOTCOD)=3"

If oconect.AbreConexion = False Then Exit Function
Set rs = oconect.CargaRecordSet(sql)
Set GetMotivosDepRet = rs
oconect.CierraConexion
Set oconect = Nothing
End Function

Public Function GrabaDepRetBancosVarios(ByVal psFormatoFecha As String, _
                                        ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                        ByVal psCtaContBanco As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                        ByVal pnDocTpo As TpoDoc, ByVal psDocNro As String, ByVal psFechaDoc As String, _
                                        ByVal psDocNroVoucher As String, ByVal rsObjHaber As ADODB.Recordset, _
                                        ByVal lnTpoObjIf As TpoObjetos, ByVal psCtaIf As String, ByVal pbDeposito As Boolean, _
                                        Optional pnMotivoNANC As Long = -1, Optional psMotObjPadre As String = "", Optional psMotObjeto As String = "", Optional ByVal psCtaAho As String = "", _
                                        Optional ByVal pnDocTpoNANCAux As TpoDoc = -1, Optional ByVal psNroDocNANCAux As String = "", Optional ByVal pnImporteAux As Currency, _
                                        Optional pnMotivoNANCAux As Long = -1, Optional psMotObjPadreAux As String = "", Optional psMotObjetoAux As String = "", Optional ByVal psCtaAhoAux As String = "") As Integer
            
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lnMovNro As Long
Dim lbTrans As Boolean
Dim oDocRec As NDocRec
Dim lsSubCta As String
Set oDMov = New DMov
Set oDocRec = New NDocRec
Dim oCont As NContFunciones
Set oCont = New NContFunciones
On Error GoTo ErrorGrabaDepRetBancosVarios

GrabaDepRetBancosVarios = 1
oDMov.Inicio psFormatoFecha
oDMov.BeginTrans
lbTrans = True

oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1: lnMovOrden = lnMovOrden + 1
lsSubCta = oCont.GetFiltroObjetos(ObjEntidadesFinancieras, psCtaContBanco, psCtaIf, False)
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaContBanco + lsSubCta, pnImporte * IIf(pbDeposito, 1, -1)
If lnTpoObjIf <> -1 Then
    Select Case lnTpoObjIf
        Case ObjCMACAgenciaArea
            lnMovOrden = lnMovOrden + 1
            oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, lnTpoObjIf
            oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(psCtaIf, 4, 2), Mid(psCtaIf, 1, 3)
        Case ObjDescomEfectivo
            lnMovOrden = lnMovOrden + 1
            oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, lnTpoObjIf
            oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, psCtaIf
        Case ObjEntidadesFinancieras
            lnMovOrden = lnMovOrden + 1
            oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
            oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psCtaIf, 4, 13), Mid(psCtaIf, 1, 2), Mid(psCtaIf, 18, 10)
            
            'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
            oDMov.InsertaMovCtaIF lnMovNro, Mid(psCtaIf, 4, 13), Mid(psCtaIf, 1, 2), Mid(psCtaIf, 18, 10), CGCtaIFConCapital, pnImporte * IIf(pbDeposito, 1, -1)
            oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                                        Mid(psCtaIf, 4, 13), Mid(psCtaIf, 1, 2), Mid(psCtaIf, 18, 10)

        Case Else
            lnMovOrden = lnMovOrden + 1
            oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, psCtaIf
    End Select
End If

lsSubCta = ""
If Not rsObjHaber Is Nothing Then
    If rsObjHaber.State = adStateOpen Then
        If Not rsObjHaber.EOF And Not rsObjHaber.BOF Then
            Do While Not rsObjHaber.EOF
                lsSubCta = lsSubCta + rsObjHaber!SubCta
                rsObjHaber.MoveNext
            Loop
        End If
    End If
End If
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber + lsSubCta, pnImporte * IIf(pbDeposito, -1, 1)
If Not rsObjHaber Is Nothing Then
    If rsObjHaber.State = adStateOpen Then
        rsObjHaber.MoveFirst
        Do While Not rsObjHaber.EOF
            lnMovOrden = lnMovOrden + 1
            Select Case rsObjHaber!cObjetoCod
                Case ObjCMACAgenciaArea
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!cObjetoCod
                    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(rsObjHaber!objeto, 4, 2), Mid(rsObjHaber!objeto, 1, 3)
                Case ObjDescomEfectivo
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!cObjetoCod
                    oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!objeto
                Case ObjEntidadesFinancieras
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!cObjetoCod
                    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(rsObjHaber!objeto, 4, 13), Mid(rsObjHaber!objeto, 1, 2), Mid(rsObjHaber!objeto, 18, 10)
                    
                    'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
                    oDMov.InsertaMovCtaIF lnMovNro, Mid(rsObjHaber!objeto, 4, 13), Mid(rsObjHaber!objeto, 1, 2), Mid(rsObjHaber!objeto, 18, 10), CGCtaIFConCapital, pnImporte * IIf(pbDeposito, -1, 1)
                    oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                                                Mid(rsObjHaber!objeto, 4, 13), Mid(rsObjHaber!objeto, 1, 2), Mid(rsObjHaber!objeto, 18, 10)
                
                Case Else
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!objeto
            End Select
            rsObjHaber.MoveNext
        Loop
    End If
End If

Select Case Val(pnDocTpo)
    Case TpoDocNotaAbono, TpoDocNotaCargo
        'oDMov.InsertaNotaAbonoCargo pnDocTpo, psDocNro, gNCNARegistrado, pnMotivoNANC, Abs(pnImporte), psMotObjPadre, psMotObjeto
        'oDMov.InsertaNotaAbonoCargoEst pnDocTpo, psDocNro, gNCNARegistrado, psMovNro
        'oDMov.InsertaRegDocCuenta pnDocTpo, psDocNro, psCtaAho
        oDMov.ActualizaNotaAbonoCargo pnDocTpo, psDocNro, gNCNAEnMovimiento
End Select
'Optional ByVal pnDocTpoNANCAux As TpoDoc = -1,  ,  Optional ByVal psNroDocNANCAux As String = "", _
Optional pnMotivoNANCAux As Long = -1, Optional psMotObjPadreAux As String = "", Optional psMotObjetoAux As String = "", Optional ByVal psCtaAhoAux As String = ""
If psDocNro <> "" Then
    oDMov.InsertaMovDoc lnMovNro, pnDocTpo, psDocNro, psFechaDoc
End If
If psDocNroVoucher <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocNroVoucher, psFechaDoc
End If
If pnDocTpoNANCAux <> -1 And psNroDocNANCAux <> "" Then
    'oDMov.InsertaNotaAbonoCargo pnDocTpoNANCAux, psNroDocNANCAux, gNCNARegistrado, pnMotivoNANCAux, pnImporteAux, psMotObjPadreAux, psMotObjetoAux
    'oDMov.InsertaNotaAbonoCargoEst pnDocTpoNANCAux, psNroDocNANCAux, gNCNARegistrado, psMovNro
    'oDMov.InsertaRegDocCuenta pnDocTpoNANCAux, psNroDocNANCAux, psCtaAhoAux
    oDMov.ActualizaNotaAbonoCargo pnDocTpoNANCAux, psNroDocNANCAux, gNCNAEnMovimiento
    oDMov.InsertaMovDoc lnMovNro, pnDocTpoNANCAux, psNroDocNANCAux, psFechaDoc
End If

GrabaDepRetBancosVarios = 0

oDMov.CommitTrans
lbTrans = False
Set oDMov = Nothing
Exit Function
ErrorGrabaDepRetBancosVarios:
    If lbTrans Then
        oDMov.RollBackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaRendicionGiroDocumento", "Error Graba Rendicion Giro de documentos"
End Function
Public Function GrabaAperturaEfectivo(ByVal psFormatoFecha As String, _
                                ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                ByVal rsBillete As ADODB.Recordset, ByVal rsMoneda As ADODB.Recordset, _
                                ByVal psCtaDebe As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                ByVal pbCreaSubCta As Boolean, ByVal psSubCtaIFCod As String, ByVal psSubCtaIFDesc As String, _
                                ByVal psPersCod As String, ByVal psIFTpo As String, _
                                ByVal psCtaIfCod As String, ByVal psCtaIFDesc As String, _
                                ByVal pdCtaIFAper As Date, ByVal psCtaIFVenc As String, _
                                ByVal pnCtaIFPlazo As Integer, ByVal pnPeriodo As Integer, ByVal pnInteres As Currency, _
                                ByVal pnTpoDocAper As TpoDoc, ByVal psNroDocAper As String, _
                                ByVal pdFechaDocApera As String) As Integer
            
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim oContF As NContFunciones
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New DMov
On Error GoTo GrabaAperturaEfectivoErr

Set oContF = New NContFunciones

GrabaAperturaEfectivo = 1
oDMov.Inicio psFormatoFecha
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable
lnMovNro = oDMov.GetnMovNro(psMovNro)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1:
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe, pnImporte
lnMovOrden = lnMovOrden + 1
If pbCreaSubCta Then
    oDMov.InsertaCuentaIF psPersCod, psIFTpo, psSubCtaIFCod, psSubCtaIFDesc, pdCtaIFAper, psCtaIFVenc, pdCtaIFAper, 0, gEstadoCtaIFActiva, psMovNro
End If
oDMov.InsertaCuentaIF psPersCod, psIFTpo, psCtaIfCod, psCtaIFDesc, pdCtaIFAper, psCtaIFVenc, pdCtaIFAper, _
                        pnCtaIFPlazo, gEstadoCtaIFRegistrada, psMovNro

oDMov.InsertaCuentaIFInteres psPersCod, psIFTpo, psCtaIfCod, pdCtaIFAper, pnPeriodo, pnInteres, lsMovNro

oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIfCod
'
''INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
'oDMov.InsertaMovCtaIF lnMovNro, psPersCod, psIFTpo, psCtaIFCod, CGCtaIFConCapital, pnImporte
'oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
'                            psPersCod, psIFTpo, psCtaIFCod
'

'guardamos el billetaje
If Not rsBillete Is Nothing Then
    If rsBillete.State = adStateOpen Then
        If Not rsBillete.EOF And Not rsBillete.BOF Then
            Do While Not rsBillete.EOF
                If rsBillete!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber, rsBillete!Monto * -1
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oDMov.InsertaMovObjEfectivo lnMovNro, Format(lnMovItem, "#0"), Format(lnMovOrden, "#0"), rsBillete!cEfectivoCod
                End If
                rsBillete.MoveNext
            Loop
        End If
        rsBillete.Close: Set rsBillete = Nothing
    End If
End If
If Not rsMoneda Is Nothing Then
    If rsMoneda.State = adStateOpen Then
        If Not rsMoneda.EOF And Not rsMoneda.BOF Then
            Do While Not rsMoneda.EOF
                If rsMoneda!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber, rsMoneda!Monto * -1
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oDMov.InsertaMovObjEfectivo lnMovNro, Format(lnMovItem, "#0"), Format(lnMovOrden, "#0"), rsMoneda!cEfectivoCod
                End If
                rsMoneda.MoveNext
            Loop
        End If
        rsMoneda.Close: Set rsMoneda = Nothing
    End If
End If
If psNroDocAper <> "" Then
    oDMov.InsertaMovDoc lnMovNro, pnTpoDocAper, psNroDocAper, pdFechaDocApera
End If

oDMov.CommitTrans
lbTrans = False

GrabaAperturaEfectivo = 0
Set oDMov = Nothing
Set oContF = Nothing
Exit Function
GrabaAperturaEfectivoErr:
    If lbTrans Then
        oDMov.RollBackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaAperturaEfectivo", Err.Description & " Error Graba Apertura de Efectivo"
End Function
Public Function GrabaAperturaDoc(ByVal psFormatoFecha As String, _
                                ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                ByVal psCtaDebe As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                ByVal pnTpoObjHaber As TpoObjetos, ByVal psObjetoCodHaber As String, _
                                ByVal pbCreaSubCta As Boolean, ByVal psSubCtaIFCod As String, ByVal psSubCtaIFDesc As String, _
                                ByVal psPersCod As String, ByVal psIFTpo As String, _
                                ByVal psCtaIfCod As String, ByVal psCtaIFDesc As String, _
                                ByVal pdCtaIFAper As Date, ByVal psCtaIFVenc As String, _
                                ByVal pnCtaIFPlazo As Integer, ByVal pnPeriodo As Integer, ByVal pnInteres As Currency, _
                                ByVal pnTpoDocAper As TpoDoc, ByVal psNroDocAper As String, _
                                ByVal pdFechaDocApera As String, _
                                ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pdFechaDoc As Date, ByVal psNroVoucher As String, _
                                Optional ByVal psCtaAhorros As String, Optional ByVal pnMotivoNANC As MotivoNotaAbonoCargo = -1, Optional ByVal psObjetoPadre As String = "", Optional psObjeto As String = "") As Integer
            
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim oContF As NContFunciones
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New DMov
On Error GoTo GrabaAperturaDocErr

Set oContF = New NContFunciones

GrabaAperturaDoc = 1
oDMov.Inicio psFormatoFecha
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable
lnMovNro = oDMov.GetnMovNro(psMovNro)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1:
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe, pnImporte
lnMovOrden = lnMovOrden + 1
If pbCreaSubCta Then
    oDMov.InsertaCuentaIF psPersCod, psIFTpo, psSubCtaIFCod, psSubCtaIFDesc, pdCtaIFAper, psCtaIFVenc, pdCtaIFAper, 0, gEstadoCtaIFActiva, psMovNro
End If
oDMov.InsertaCuentaIF psPersCod, psIFTpo, psCtaIfCod, psCtaIFDesc, pdCtaIFAper, psCtaIFVenc, pdCtaIFAper, _
                        pnCtaIFPlazo, gEstadoCtaIFRegistrada, psMovNro

oDMov.InsertaCuentaIFInteres psPersCod, psIFTpo, psCtaIfCod, pdCtaIFAper, pnPeriodo, pnInteres, psMovNro

oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIfCod

'
''INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
'oDMov.InsertaMovCtaIF lnMovNro, psPersCod, psIFTpo, psCtaIFCod, CGCtaIFConCapital, pnImporte
'oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
'                            psPersCod, psIFTpo, psCtaIFCod


lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber, pnImporte * -1
lnMovOrden = lnMovOrden + 1
If pnTpoObjHaber <> -1 Then
Select Case pnTpoObjHaber
    Case ObjEntidadesFinancieras
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, pnTpoObjHaber
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psObjetoCodHaber, 4, 13), Mid(psObjetoCodHaber, 1, 2), Mid(psObjetoCodHaber, 18, 10)
        
        'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
        oDMov.InsertaMovCtaIF lnMovNro, Mid(psObjetoCodHaber, 4, 13), Mid(psObjetoCodHaber, 1, 2), Mid(psObjetoCodHaber, 18, 10), CGCtaIFConCapital, pnImporte * -1
        oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                                    Mid(psObjetoCodHaber, 4, 13), Mid(psObjetoCodHaber, 1, 2), Mid(psObjetoCodHaber, 18, 10)
    
    Case ObjCMACAgenciaArea
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, pnTpoObjHaber
        oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(psObjetoCodHaber, 4, 2), Mid(psObjetoCodHaber, 1, 3)
    Case ObjDescomEfectivo
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, pnTpoObjHaber
        oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, psObjetoCodHaber
    Case Else
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, psObjetoCodHaber
End Select
End If
If pnTpoDoc = TpoDocNotaAbono Or pnTpoDoc = TpoDocNotaCargo Then
    oDMov.ActualizaNotaAbonoCargo pnTpoDoc, psNroDoc, gNCNAEnMovimiento, , pnImporte
    oDMov.InsertaNotaAbonoCargoEst pnTpoDoc, psNroDoc, gNCNAEnMovimiento, lsMovNro
End If
If psNroDoc <> "" Then
    oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, pdFechaDoc
End If
If psNroVoucher <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psNroVoucher, pdFechaDoc
End If

If psNroDocAper <> "" Then
    oDMov.InsertaMovDoc lnMovNro, pnTpoDocAper, psNroDocAper, pdFechaDocApera
End If

oDMov.CommitTrans
lbTrans = False

GrabaAperturaDoc = 0
Set oDMov = Nothing
Set oContF = Nothing
Exit Function
GrabaAperturaDocErr:
    If lbTrans Then
        oDMov.RollBackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaAperturaDoc", Err.Description & " Error Graba Apertura con Trasferencia"
End Function
Public Function GrabaAperturaRegCheque(ByVal psFormatoFecha As String, ByVal psMovNro As String, ByVal psOpeCod As String, _
                                      ByVal psMovDesc As String, _
                                      ByVal psCtaDebe As String, ByVal psObjCodProd As String, ByVal psAreaCod As String, ByVal psAgeCod As String, _
                                      ByVal psCtaHaber As String, ByVal rsObjHaber As ADODB.Recordset, _
                                      ByVal psNumCheque As String, ByVal psPersCodIf As String, ByVal psTipoIF As String, ByVal pnPlaza As ChequePlaza, _
                                      ByVal psCtaCheque As String, ByVal pnImporte As Currency, ByVal pdFechaReg As Date, _
                                      ByVal pdFechaVal As Date, ByVal pnMonedaCheque As Moneda, _
                                      ByVal pbCreaSubCta As Boolean, ByVal psSubCtaIFCod As String, ByVal psSubCtaIFDesc As String, _
                                      ByVal psPersCod As String, ByVal psIFTpo As String, _
                                      ByVal psCtaIfCod As String, ByVal psCtaIFDesc As String, _
                                      ByVal pdCtaIFAper As Date, ByVal psCtaIFVenc As String, _
                                      ByVal pnCtaIFPlazo As Integer, ByVal pnPeriodo As Integer, ByVal pnInteres As Currency, _
                                      ByVal pnTpoDocAper As TpoDoc, ByVal psNroDocAper As String, _
                                      ByVal pdFechaDocApera As String, _
                                      Optional ByVal pnEstadoCheque As ChequeEstado = gChqEstEnValorizacion, _
                                      Optional ByVal pnEstadoCG As CGEstadosChq = gCGEstadosChqRecibido, _
                                      Optional ByVal pnConfirmaCaja As CGEstadoConfCheque = ChqCGSinConfirmacion, _
                                      Optional ByVal psAreaCodChq As String, Optional ByVal psAgeCodChq As String) As Integer

Dim oMov As DMov
Dim lsSubCuenta As String
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New DMov

On Error GoTo GrabaAperturaRegChequeErr
lsSubCuenta = ""
oMov.Inicio psFormatoFecha
oMov.BeginTrans
lbTrans = True
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable
lnMovNro = oMov.GetnMovNro(psMovNro)
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1
oMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe, pnImporte
lnMovOrden = lnMovOrden + 1

If pbCreaSubCta Then
    oMov.InsertaCuentaIF psPersCod, psIFTpo, psSubCtaIFCod, psSubCtaIFDesc, pdCtaIFAper, psCtaIFVenc, pdCtaIFAper, 0, gEstadoCtaIFActiva, psMovNro
End If
oMov.InsertaCuentaIF psPersCod, psIFTpo, psCtaIfCod, psCtaIFDesc, pdCtaIFAper, psCtaIFVenc, pdCtaIFAper, _
                        pnCtaIFPlazo, gEstadoCtaIFRegistrada, psMovNro

oMov.InsertaCuentaIFInteres psPersCod, psIFTpo, psCtaIfCod, pdCtaIFAper, pnPeriodo, pnInteres, psMovNro

oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIfCod

'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
'oDMov.InsertaMovCtaIF lnMovNro, psPersCod, psIFTpo, psCtaIFCod, CGCtaIFConCapital, pnImporte
'oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                            psPersCod, psIFTpo, psCtaIFCod

lnMovItem = lnMovItem + 1: lnMovOrden = 0
oMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber, pnImporte * -1
oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, psObjCodProd
lnMovOrden = lnMovOrden + 1
oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
oMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, psAgeCod, psAreaCod

If psNumCheque <> "" Then
    oMov.InsertaMovDoc lnMovNro, TpoDocCheque, psNumCheque, pdFechaReg
End If
'grabacion dentro de la tabla de cheques
oMov.InsertaCheque TpoDocCheque, psNumCheque, psPersCodIf, psTipoIF, pnPlaza, psCtaCheque, pnImporte, pdFechaReg, _
                    pdFechaVal, pnEstadoCG, pnConfirmaCaja, pnMonedaCheque, , psAreaCodChq, psAgeCodChq

oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIf, psTipoIF, pdFechaReg, gChqEstEnValorizacion, psMovNro

If psNroDocAper <> "" Then
    oMov.InsertaMovDoc lnMovNro, pnTpoDocAper, psNroDocAper, pdFechaDocApera
End If

oMov.CommitTrans
lbTrans = False
GrabaAperturaRegCheque = 0
Exit Function
GrabaAperturaRegChequeErr:
    If lbTrans Then
        oMov.RollBackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetApertSinConf(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset
Dim sql As String
Dim oconect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltroAge As String
Set rs = New ADODB.Recordset
Set oconect = New DConecta
If oconect.AbreConexion = False Then Exit Function


sql = "Select   CONVERT(CHAR(10),CONVERT(DATETIME, SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHA," _
    & "         P.CPERSNOMBRE AS BANCO, SC.cCtaIFDesc + ' ' + C.cCtaIFDesc AS CUENTA, " _
    & "         CONVERT(CHAR(10),C.dCtaIFAper,103) AS APERTURA, ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)  AS IMPORTE,  M.NMOVNRO, M.COPECOD, C.CCTAIFCOD, M.CMOVDESC, C.cIFTpo, C.cPersCod  " _
    & " From    Mov M " _
    & "         JOIN MovCta MC ON MC.nMovNro =M.nMovNro " _
    & "         LEFT JOIN MOVME ME ON ME.NMOVNRO=MC.NMOVNRO " _
    & "         JOIN MOVOBJ MO on MO.nMovNro = MC.nMovNro and MO.nmovitem =MC.NMOVITEM " _
    & "         JOIN MOVOBJIF MIF ON MIF.nMovNro = MO.nMovNro and MIF.nMovItem= MO.NMOVITEM AND MO.NMOVOBJORDEN = MIF.NMOVOBJORDEN " _
    & "         JOIN CTAIF C ON C.cPersCod =MIF.cPersCod AND C.cIFTpo = MIF.cIFTpo AND C.cCtaIfCod = MIF.cCtaIfCod " _
    & "         JOIN CTAIF SC ON SC.cPersCod =MIF.cPersCod AND SC.cIFTpo = MIF.cIFTpo AND SUBSTRING(MIF.cCtaIfCod,1,LEN(SC.cCtaIfCod)) = SC.cCtaIfCod AND LEN(SC.cCtaIfCod)<=5 " _
    & "         JOIN PERSONA P ON P.CPERSCOD = C.CPERSCOD " _
    & "         JOIN OPECTA OC ON OC.cCtaContCod= MC.cCtaContCod " _
    & "         JOIN OPEOBJ OO ON OO.COPECOD = OC.COPECOD  " _
    & " Where   OC.COPECOD ='" & psOpeCod & "' AND cOpeCtaDH='H' AND MC.NMOVIMPORTE>0 AND M.NMOVFLAG =" & gMovFlagVigente & "" _
    & "         AND C.cIFTpo + C.CCTAIFCOD LIKE cOpeObjFiltro  " _
    & "         AND SUBSTRING(M.cMovnro,1,8 ) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "'  " _
    & "         AND NOT EXISTS (SELECT  M1.NMOVNRO " _
    & "                         FROM    MOV M1 JOIN MOVREF MR ON MR.NMOVNRO = M1.NMOVNRO " _
    & "                         WHERE   MR.NMOVNROREF =M.NMOVNRO AND M1.NMOVFLAG =" & gMovFlagVigente & ")"
    
Set rs = oconect.CargaRecordSet(sql)
Set GetApertSinConf = rs
oconect.CierraConexion
Set oconect = Nothing

End Function
Public Function GrabaConfApertura(ByVal psFormatoFecha As String, _
                                ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                ByVal psCtaDebe As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psCtaIfCod As String, _
                                ByVal pnMovNroRef As Long, ByVal pdFecha As Date) As Integer
           
          
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim oCtaIf As NCajaCtaIF
Dim lnSaldo As Currency
Set oCtaIf = New NCajaCtaIF
Set oDMov = New DMov
On Error GoTo GrabaConfAperturaErr


GrabaConfApertura = 1
oDMov.Inicio psFormatoFecha
lsSubCta = oCtaIf.GeneraSubCuentaIF(psPersCod, psIFTpo, psCtaDebe, psCtaIfCod)
If lsSubCta = "" Then
    Err.Raise 10000, "", "No se ha podido determinar sub cuenta de Cuenta Seleccionada"
    Exit Function
End If

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc
lnMovNro = oDMov.GetnMovNro(psMovNro)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1:
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe + lsSubCta, pnImporte
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIfCod

'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
oDMov.InsertaMovCtaIF lnMovNro, psPersCod, psIFTpo, psCtaIfCod, CGCtaIFConCapital, pnImporte
oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                            psPersCod, psIFTpo, psCtaIfCod

oDMov.InsertaCuentaIFFiltro psPersCod, psIFTpo, psCtaIfCod, psCtaDebe, lsSubCta

lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber, pnImporte * -1
oDMov.ActualizaMov pnMovNroRef, , gMovEstContabMovContable

'Actualizamos el estado del movimiento de registro apertura en un movimiento Contable
oDMov.ActualizaMov pnMovNroRef, , gMovEstContabMovContable
'pasamos a activo el estado de la cuenta aperturada
oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIfCod, , , , , , gEstadoCtaIFActiva

oDMov.InsertaMovRef lnMovNro, pnMovNroRef

oDMov.CommitTrans
lbTrans = False

GrabaConfApertura = 0
Set oDMov = Nothing

Exit Function
GrabaConfAperturaErr:
    If lbTrans Then
        oDMov.RollBackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaConfApertura", Err.Description & " Error en Grabación de Confirmación de Apertura "
End Function
Public Function GrabaInteresDev(ByVal psFormatoFecha As String, _
                                ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                ByVal psCtaDebe As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                ByVal psPersCod As String, ByVal psIFTpo As String, _
                                ByVal psCtaIfCod As String, ByVal pdCtaIFInt As Date, ByVal pnDiasTrans As Integer) As Integer

Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lnInteres As Currency
Dim oCtaIf As NCajaCtaIF
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New DMov
On Error GoTo GrabaInteresDevErr

GrabaInteresDev = 1
Set oCtaIf = New NCajaCtaIF
lnInteres = oCtaIf.GetInteresCtaIf(psPersCod, psIFTpo, psCtaIfCod)
lnInteres = lnInteres + pnImporte
Set oCtaIf = Nothing

oDMov.Inicio psFormatoFecha
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc
lnMovNro = oDMov.GetnMovNro(psMovNro)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1:
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe, pnImporte
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIfCod

''INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
oDMov.InsertaMovCtaIF lnMovNro, psPersCod, psIFTpo, psCtaIfCod, CGCtaIFConInteres, pnImporte, pnDiasTrans
'oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                            psPersCod, psIFTpo, psCtaIFCod

lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber, pnImporte * -1

oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIfCod, , , , pdCtaIFInt, , , lnInteres, psMovNro

oDMov.CommitTrans
lbTrans = False

GrabaInteresDev = 0
Set oDMov = Nothing
Set oCtaIf = Nothing
Exit Function
GrabaInteresDevErr:
    If lbTrans Then
        oDMov.RollBackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaInteresDev", Err.Description & " Error Grabación de Interes Devengado"
End Function
Public Function GrabaCapitalizaRegCheque(ByVal psFormatoFecha As String, ByVal psMovNro As String, ByVal psOpeCod As String, _
                                        ByVal psMovDesc As String, ByVal pdFechaMov As Date, _
                                        ByVal psCtaDebe As String, ByVal psObjCodProd As String, ByVal psAreaCod As String, ByVal psAgeCod As String, _
                                        ByVal psCtaHaberInteres As String, ByVal psCtaHaberIntCalc As String, _
                                        ByVal psCtaHaberCapital As String, ByVal pnImporteCapital As Currency, _
                                        ByVal pnImporteInt As Currency, ByVal pnImporteIntCalc As Currency, _
                                        ByVal rsObjHaberMot As ADODB.Recordset, _
                                        ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, ByVal psCtaIfCod As String, _
                                        ByVal psNumCheque As String, ByVal psPersCodIfChq As String, ByVal pnTipoIFChq As CGTipoIF, ByVal pnPlaza As ChequePlaza, _
                                        ByVal psCtaCheque As String, ByVal pnImporteTotal As Currency, ByVal pdFechaReg As Date, _
                                        ByVal pdFechaVal As Date, ByVal pnMonedaCheque As Moneda, _
                                        Optional ByVal pnEstadoCheque As ChequeEstado = gChqEstEnValorizacion, _
                                        Optional ByVal pnEstadoCG As CGEstadosChq = gCGEstadosChqRecibido, _
                                        Optional ByVal pnConfirmaCaja As CGEstadoConfCheque = ChqCGSinConfirmacion, _
                                        Optional ByVal psAreaCodChq As String, Optional ByVal psAgeCodChq As String, Optional pnDiasTrans As Integer = 0) As Integer

Dim oMov As DMov
Dim lsSubCuenta As String
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New DMov

On Error GoTo GrabaCapitalizaRegChequeErr

GrabaCapitalizaRegCheque = 1
lsSubCuenta = ""
oMov.Inicio psFormatoFecha
oMov.BeginTrans
lbTrans = True
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc
lnMovNro = oMov.GetnMovNro(psMovNro)
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1
oMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe, pnImporteTotal
lnMovOrden = lnMovOrden + 1
oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, psObjCodProd
lnMovOrden = lnMovOrden + 1
oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
oMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, psAgeCod, psAreaCod

If pnImporteInt <> 0 Then
    lnMovItem = lnMovItem + 1: lnMovOrden = 0
    oMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaberInteres, pnImporteInt * -1
End If
If pnImporteIntCalc <> 0 Then
    lnMovItem = lnMovItem + 1: lnMovOrden = 0
    oMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaberIntCalc, pnImporteIntCalc * -1
End If

'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
oMov.InsertaMovCtaIF lnMovNro, psPersCod, pnIFTpo, psCtaIfCod, CGCtaIFConInteres, (pnImporteIntCalc + pnImporteInt) * -1, pnDiasTrans

If pnImporteCapital > 0 And psCtaHaberCapital <> "" Then
    lnMovItem = lnMovItem + 1: lnMovOrden = 0
    oMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaberCapital, pnImporteCapital * -1
    
    'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
    oMov.InsertaMovCtaIF lnMovNro, psPersCod, pnIFTpo, psCtaIfCod, CGCtaIFConCapital, pnImporteCapital * -1
    oMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                            psPersCod, pnIFTpo, psCtaIfCod
                            
End If

InsertaRefCap psPersCod, pnIFTpo, psCtaIfCod, lnMovNro, oMov
oMov.ActualizaCtaIF psPersCod, pnIFTpo, psCtaIfCod, , , pdFechaMov, , , , 0, psMovNro
'grabacion dentro de la tabla de cheques
oMov.InsertaCheque TpoDocCheque, psNumCheque, psPersCodIfChq, pnTipoIFChq, pnPlaza, psCtaCheque, pnImporteTotal, pdFechaReg, _
                    pdFechaVal, pnEstadoCG, pnConfirmaCaja, pnMonedaCheque, , psAreaCodChq, psAgeCodChq
oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIfChq, pnTipoIFChq, pdFechaReg, gChqEstEnValorizacion, psMovNro

If psNumCheque <> "" Then
    oMov.InsertaMovDoc lnMovNro, TpoDocCheque, psNumCheque, pdFechaReg
End If

oMov.CommitTrans
lbTrans = False
GrabaCapitalizaRegCheque = 0
Exit Function
GrabaCapitalizaRegChequeErr:
    If lbTrans Then
        oMov.RollBackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Sub InsertaRefCap(ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, _
                          ByVal psCtaIfCod As String, ByVal pnMovNro As Long, ByRef oMov As DMov)
Dim rs As ADODB.Recordset
Dim sql As String

sql = " SELECT  DISTINCT M.NMOVNRO, M.CMOVNRO , MC.NMOVIMPORTE" _
    & " FROM    MOV M " _
    & "         JOIN MOVCTAIF MC ON MC.NMOVNRO = M.NMOVNRO " _
    & "         JOIN CTAIF C     ON MC.CPERSCOD=C.CPERSCOD AND MC.CIFTPO = C.CIFTPO AND MC.cCtaIFCod=C.cCtaIFCod " _
    & " Where   MC.nConceptoCod =" & CGCtaIFConInteres & " And MC.NMOVIMPORTE > 0 And m.nMovFlag = " & gMovFlagVigente & " " _
    & "         AND SUBSTRING(M.CMOVNRO,1,8)> CONVERT(CHAR(8),dCtaIfCap,112) " _
    & "         AND MC.cPersCod='" & psPersCod & "' AND MC.cIFTpo='" & Format(pnIFTpo, "00") & "' AND MC.cCtaIFCod ='" & psCtaIfCod & "'"

Set rs = oMov.CargaRecordSet(sql)
Do While Not rs.EOF
    oMov.InsertaMovRef pnMovNro, rs!nMovNro
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
End Sub
Public Function GrabaCapitalizaIFEfectivo(ByVal psFormatoFecha As String, _
                                         ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                         ByVal rsBillete As ADODB.Recordset, ByVal rsMoneda As ADODB.Recordset, _
                                         ByVal psCtaDebe As String, ByVal pnImporte As Currency, _
                                         ByVal psCtaDif As String, ByVal pnImporteDif As Currency, _
                                         ByVal psCtaHaberInt As String, ByVal pnImporteInt As Currency, _
                                         ByVal psCtaHaberIntCal As String, ByVal pnImporteIntCalc As Currency, _
                                         ByVal psCtaHaberCapital As String, ByVal pnImporteCapital As Currency, _
                                         ByVal psPersCod As String, ByVal psIFTpo As String, _
                                         ByVal psCtaIfCod As String, ByVal pdCtaIFCap As Date, Optional pnDiasTrans As Integer = 0) As Integer
                    
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim oContF As NContFunciones
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New DMov
On Error GoTo GrabaCapitalizaIFEfectivoErr

Set oContF = New NContFunciones

GrabaCapitalizaIFEfectivo = 1
oDMov.Inicio psFormatoFecha
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc
lnMovNro = oDMov.GetnMovNro(psMovNro)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
'guardamos el billetaje
If Not rsBillete Is Nothing Then
    If rsBillete.State = adStateOpen Then
        If Not rsBillete.EOF And Not rsBillete.BOF Then
            Do While Not rsBillete.EOF
                If rsBillete!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe, rsBillete!Monto
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oDMov.InsertaMovObjEfectivo lnMovNro, Format(lnMovItem, "#0"), Format(lnMovOrden, "#0"), rsBillete!cEfectivoCod
                End If
                rsBillete.MoveNext
            Loop
        End If
        rsBillete.Close: Set rsBillete = Nothing
    End If
End If
If Not rsMoneda Is Nothing Then
    If rsMoneda.State = adStateOpen Then
        If Not rsMoneda.EOF And Not rsMoneda.BOF Then
            Do While Not rsMoneda.EOF
                If rsMoneda!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe, rsMoneda!Monto
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oDMov.InsertaMovObjEfectivo lnMovNro, Format(lnMovItem, "#0"), Format(lnMovOrden, "#0"), rsMoneda!cEfectivoCod
                End If
                rsMoneda.MoveNext
            Loop
        End If
        rsMoneda.Close: Set rsMoneda = Nothing
    End If
End If
If pnImporteDif > 0 Then
    lnMovItem = lnMovItem + 1:
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDif, pnImporteDif
End If

If pnImporteInt > 0 Then
    lnMovItem = lnMovItem + 1:
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaberInt, pnImporteInt * -1
End If
If pnImporteIntCalc > 0 Then
    lnMovItem = lnMovItem + 1:
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaberIntCal, pnImporteIntCalc * -1
End If
''INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
oDMov.InsertaMovCtaIF lnMovNro, psPersCod, psIFTpo, psCtaIfCod, CGCtaIFConInteres, (pnImporteIntCalc + pnImporteInt) * -1, pnDiasTrans

InsertaRefCap psPersCod, psIFTpo, psCtaIfCod, lnMovNro, oDMov
oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIfCod, , , pdCtaIFCap, , , , 0, psMovNro

If pnImporteCapital > 0 And psCtaHaberCapital <> "" Then
    lnMovItem = lnMovItem + 1:
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaberCapital, pnImporteCapital * -1
    
    'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
    oDMov.InsertaMovCtaIF lnMovNro, psPersCod, psIFTpo, psCtaIfCod, CGCtaIFConCapital, pnImporteCapital * -1
    oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                            psPersCod, psIFTpo, psCtaIfCod
    
    oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIfCod, , , pdCtaIFCap, , , gEstadoCtaIFCancelada, 0, psMovNro
End If

oDMov.CommitTrans
lbTrans = False

GrabaCapitalizaIFEfectivo = 0
Set oDMov = Nothing
Set oContF = Nothing
Exit Function
GrabaCapitalizaIFEfectivoErr:
    If lbTrans Then
        oDMov.RollBackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaCapitalizaIFEfectivo", Err.Description & " Error Graba Operaciones en Efectivo de Cuentas de IF"
End Function
Public Function GrabaCapitalizaDoc(ByVal psFormatoFecha As String, _
                                    ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                    ByVal psCtaDebe As String, ByVal pnImporte As Currency, _
                                    ByVal pnTpoObj As TpoObjetos, ByVal psObjetoCodDebe As String, _
                                    ByVal psCtaHaberInt As String, ByVal pnImporteInt As Currency, _
                                    ByVal psCtaHaberIntCal As String, ByVal pnImporteIntCalc As Currency, _
                                    ByVal psCtaHaberCapital As String, ByVal pnImporteCapital As Currency, _
                                    ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psCtaIfCod As String, _
                                    ByVal pdCtaIFCap As Date, _
                                    ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pdFechaDoc As Date, ByVal psNroVoucher As String, _
                                    Optional ByVal psCtaAhorros As String, Optional ByVal pnMotivoNANC As MotivoNotaAbonoCargo = -1, Optional ByVal psObjetoPadre As String = "", Optional psObjeto As String = "", Optional pnDiasTrans As Integer = 0) As Integer
            
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New DMov
On Error GoTo GrabaCapitalizaDocErr



GrabaCapitalizaDoc = 1
oDMov.Inicio psFormatoFecha
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc
lnMovNro = oDMov.GetnMovNro(psMovNro)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe, pnImporte
lnMovOrden = lnMovOrden + 1
If pnTpoObj <> -1 Then
    Select Case pnTpoObj
        Case ObjEntidadesFinancieras
            oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, pnTpoObj
            oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psObjetoCodDebe, 4, 13), Mid(psObjetoCodDebe, 1, 2), Mid(psObjetoCodDebe, 18, 10)
            
            'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
            oDMov.InsertaMovCtaIF lnMovNro, Mid(psObjetoCodDebe, 4, 13), Mid(psObjetoCodDebe, 1, 2), Mid(psObjetoCodDebe, 18, 10), CGCtaIFConCapital, pnImporte
            oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                                        Mid(psObjetoCodDebe, 4, 13), Mid(psObjetoCodDebe, 1, 2), Mid(psObjetoCodDebe, 18, 10)
        Case ObjCMACAgenciaArea
            oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, pnTpoObj
            oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(psObjetoCodDebe, 4, 2), Mid(psObjetoCodDebe, 1, 3)
        Case ObjDescomEfectivo
            oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, pnTpoObj
            oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, psObjetoCodDebe
        Case Else
            oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, psObjetoCodDebe
    End Select
End If

If pnImporteInt > 0 Then
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaberInt, pnImporteInt * -1
End If
If pnImporteIntCalc > 0 Then
    lnMovItem = lnMovItem + 1:
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaberIntCal, pnImporteIntCalc * -1
End If
''INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
oDMov.InsertaMovCtaIF lnMovNro, psPersCod, psIFTpo, psCtaIfCod, CGCtaIFConInteres, (pnImporteInt + pnImporteIntCalc) * -1, pnDiasTrans

InsertaRefCap psPersCod, psIFTpo, psCtaIfCod, lnMovNro, oDMov
oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIfCod, , , pdCtaIFCap, , , , 0, psMovNro
If pnImporteCapital > 0 And psCtaHaberCapital <> "" Then
    lnMovItem = lnMovItem + 1:
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaberCapital, pnImporteCapital * -1
    
    'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
    oDMov.InsertaMovCtaIF lnMovNro, psPersCod, psIFTpo, psCtaIfCod, CGCtaIFConCapital, pnImporteCapital * -1
    oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                            psPersCod, psIFTpo, psCtaIfCod
End If

If pnTpoDoc = TpoDocNotaAbono Or pnTpoDoc = TpoDocNotaCargo Then
    oDMov.ActualizaNotaAbonoCargo pnTpoDoc, psNroDoc, gNCNAEnMovimiento, , pnImporte
    oDMov.InsertaNotaAbonoCargoEst pnTpoDoc, psNroDoc, gNCNAEnMovimiento, psMovNro
End If
If psNroDoc <> "" Then
    oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, pdFechaDoc
End If
If psNroVoucher <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psNroVoucher, pdFechaDoc
End If

oDMov.CommitTrans
lbTrans = False

GrabaCapitalizaDoc = 0
Set oDMov = Nothing
Exit Function
GrabaCapitalizaDocErr:
    If lbTrans Then
        oDMov.RollBackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaCapitalizaDoc", Err.Description & " Error Graba Capitalización con Documentos"
End Function

