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
Public Enum MotivosDepRet
    gMotDepositos = 1
    gMotRetiros = 2
End Enum

Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsMsgErr As String
Dim lbConexionPropia As Boolean
'Jeom Uso de Matriz por performance
Private pMatBillete() As String
Private pMatMoneda() As String
Private pMatMovCta() As String
Private pMatMovME() As String
Private pMatActSaldMov1() As String
Private pMatActSaldMov2() As String
Private pMatActSaldMov3() As String


Public Function GetNroCodOp() As String
GetNroCodOp = ""
sql = "Select Isnull(Max(cRangoCod),'0') as MaxCodRango From OpCajaGen"
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetNroCodOp = Format(Val(rs!maxCodRango) + 1, "0000")
End If
'RSClose Rs
rs.Close
End Function

Public Function GetEstadosOp(Optional pbFiltro As Boolean, Optional pnbEstOpDesde As CGEstadosOp = PorEmitir, Optional pnbEstOpHasta As CGEstadosOp = Extraviadas) As ADODB.Recordset
sql = " Select  cConsDescripcion, nConsValor " _
    & " From    Constante " _
    & " Where   nCONSCOD= '" & gCGEstadosOp & "' and nCONSCOD<>nConsValor "
Set rs = oConect.CargaRecordSet(sql)
If pbFiltro Then
     rs.Filter = "nConsValor >='" & pnbEstOpDesde & "'  AND nConsValor<= '" & pnbEstOpHasta & "'"
End If
Set GetEstadosOp = rs
GetEstadosOp.MoveFirst
End Function

Public Function GrabaRangoOpCajaGeneral(ByVal psRangoCod As String, ByVal pdFechaIng As Date, ByVal pnRangoEmiIni As Long, ByVal pnRangoEmiFin As Long, ByVal psEstadoRango As CGEstadosOp, ByVal psUltActualizacion As String, ByVal pnMoneda As Moneda, Optional pbEjectBacth As Boolean = False) As Integer
Dim oCajaGen As DCajaGeneral
GrabaRangoOpCajaGeneral = 1
Set oCajaGen = New DCajaGeneral
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
sql = "Select   cRangoCod, dFechaIng, nRangoEmiIni, nRangoEmiFin, " _
    & "         cEstadoRango, C.cConsDescripcion,  " _
    & "         Op.cMoneda, C1.cConsDescripcion as cConsDescMon, cUltActualizacion " _
    & " From    OPCajaGen op JOIN Constante C on C.nConsValor= Op.cEstadoRango " _
    & "         JOIN Constante C1 on C1.nConsValor= Op.cMoneda " _
    & " where   C.nCONSCOD= '" & gCGEstadosOp & "' and C1.nCONSCOD='" & gMoneda & "' ORDER BY OP.cRangoCod "
    
Set rs = oConect.CargaRecordSet(sql)
Set GetOpCajaGeneral = rs
End Function
Public Function GetOpDetCajaGeneral(ByVal psRangoCod As String) As ADODB.Recordset
sql = "Select cRangoCod, nRangoIni,nRangoFin, cEstadoDet, cConsDescripcion, cDescDet" _
    & " From OPCajaGenDet op JOIN Constante C on C.nConsValor= Op.cEstadoDet " _
    & " where C.nCONSCOD= '" & gCGEstadosOp & "' and OP.cRangoCod ='" & psRangoCod & "'"
Set rs = oConect.CargaRecordSet(sql)
Set GetOpDetCajaGeneral = rs

End Function
Public Function VerificaOPRango(ByVal psDocNro As String, ByVal psMoneda As Moneda) As String
'PRIMERO VERIFICAMOS SI EL DOCUMENTO SE ENCUENTRA DENTRO DEL RANGO ESTABLECIDO
sql = "     SELECT  OPG.cRangoCod,nRangoEmiIni,nRangoEmiFin,cMoneda," _
        & "         nRangoIni As RangoInIDet, nRangoFin As RangoFinDet, cEstadoDet " _
        & " From    OPCajaGen OPG LEFT JOIN OPCajaGenDet OPD ON OPD.cRangoCod= OPG.cRangoCod " _
        & " WHERE   OPG.CMONEDA ='" & psMoneda & "' AND OPG.cEstadoRango='" & PorEmitir & "' " _
        & "         AND " & psDocNro & " BETWEEN OPG.nRangoEmiIni AND OPG.nRangoEmiFin "

Set rs = oConect.CargaRecordSet(sql)
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
        
        Set rs = oConect.CargaRecordSet(sql)
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
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        VerificaOPRango = Format(rs!nRangoEmiIni, String(8, "0"))
    Else
        'Debera ingresar un nuevo rango de ordenes de pago listas para emitir
        VerificaOPRango = ""
    End If
End If
rs.Close
Set rs = Nothing
End Function
Public Function GrabaHabEfectivo(ByVal psMovNroHab As String, _
                                ByVal rsBillete As ADODB.Recordset, _
                                ByVal rsMoneda As ADODB.Recordset, _
                                ByVal psOpecod As String, ByVal psMovDesc As String, _
                                ByVal psCtaDebe As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                ByVal psAreaOrig As String, ByVal psAreaDest As String, _
                                ByVal psNroDocTransp As String, ByVal psFechaDocTransp As String, _
                                ByVal psNroDocSal As String, ByVal psFechaDocSal As String, _
                                ByVal psCodTransp As String, ByVal pnMontoTransp As Currency, _
                                Optional ByVal psMovNroRef As String = "", Optional ByVal psAreaCodCaja As String = "") As Integer
                                
            
Dim lnMovItem As Long
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim oContF As NContFunciones
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lnFactor As Integer
Dim lsCuenta As String

Set oDMov = New DMov
'On Error GoTo GrabaHabEfectivoErr

Set oContF = New NContFunciones

GrabaHabEfectivo = 1

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNroHab, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNroHab)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
If psOpecod = gOpeBoveCGConfHabAgeBoveMN Or psOpecod = gOpeBoveCGConfHabAgeBoveME Then
    lnMovItem = 0: lnMovOrden = 0
    'Guardamos el billetaje
    oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsBillete, psCtaDebe, "D"
    oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsMoneda, psCtaDebe, "D"
    lnMovOrden = 2
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(psAreaDest, 4, 2), Mid(psAreaDest, 1, 3)
    
    lnMovOrden = lnMovOrden + 1
    lnMovItem = lnMovItem + 1:
    'lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaHaber, psAreaOrig, False)
    lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgencias, psCtaHaber, psAreaOrig, False)
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber + lsSubCta, pnImporte * -1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(psAreaOrig, 4, 2), Mid(psAreaOrig, 1, 3)
    
Else
    lnMovItem = 0: lnMovOrden = 0
    lnMovItem = lnMovItem + 1:
    'lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgencias, psCtaDebe, psAreaDest, False)
    lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaDebe, psAreaDest, False)
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe + lsSubCta, pnImporte
    lnMovOrden = lnMovOrden + 1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(psAreaDest, 4, 2), Mid(psAreaDest, 1, 3)
    'Guardamos el billetaje
    oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsBillete, psCtaHaber, "H"
    oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsMoneda, psCtaHaber, "H"
    
    lnMovOrden = 2
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(psAreaOrig, 4, 2), Mid(psAreaOrig, 1, 3)
End If

If psNroDocSal <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocSalidaEfectivoBoveda, psNroDocSal, psFechaDocSal
End If
If psNroDocTransp <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocCompServTranspValores, psNroDocTransp, psFechaDocTransp
End If
'If psCodTransp <> "" Then
    If psAreaCodCaja = "" Then
        oDMov.InsertaMovHabilitaciones lnMovNro, psCodTransp, Mid(psAreaOrig, 1, 3), Mid(psAreaOrig, 4, 2), Mid(psAreaDest, 1, 3), Mid(psAreaDest, 4, 2), pnMontoTransp, pnImporte, Mid(psOpecod, 3, 1)
    Else
        oDMov.InsertaMovHabilitaciones lnMovNro, psCodTransp, Mid(psAreaOrig, 1, 3), Mid(psAreaOrig, 4, 2), Mid(psAreaCodCaja, 1, 3), Mid(psAreaCodCaja, 4, 2), pnMontoTransp, pnImporte, Mid(psOpecod, 3, 1)
    End If
'End If
If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovME lnMovNro, psMovNroHab
End If

oDMov.ActualizaSaldoMovimiento psMovNroHab, "+"
oDMov.CommitTrans
lbTrans = False

GrabaHabEfectivo = 0
Set oDMov = Nothing
Set oContF = Nothing
Exit Function
GrabaHabEfectivoErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaHabEfectivo", Err.Description & " Error Graba Habilitación de Efectivo"
End Function

Public Function GrabaHabEfectivoNew(ByVal psMovNroHab As String, _
                                ByVal rsBillete As ADODB.Recordset, _
                                ByVal rsMoneda As ADODB.Recordset, _
                                ByVal psOpecod As String, ByVal psMovDesc As String, _
                                ByVal psCtaDebe As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                ByVal psAreaOrig As String, ByVal psAreaDest As String, _
                                ByVal psNroDocTransp As String, ByVal psFechaDocTransp As String, _
                                ByVal psNroDocSal As String, ByVal psFechaDocSal As String, _
                                ByVal psCodTransp As String, ByVal pnMontoTransp As Currency, _
                                Optional ByVal psMovNroRef As String = "", Optional ByVal psAreaCodCaja As String = "", Optional psFechaMov As Date) As Integer
                                
            
Dim lnMovItem As Long
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim oContF As NContFunciones
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lnFactor As Integer
Dim lsCuenta As String
Dim nTCFijo As Currency
Dim nTCVenta As Currency
Dim nTCComp As Currency
Dim dFecha As String


Set oDMov = New DMov
Set oContF = New NContFunciones

GrabaHabEfectivoNew = 1
'---------------------------------------------------------------------------------
If psOpecod = gOpeBoveCGConfHabAgeBoveMN Or psOpecod = gOpeBoveCGConfHabAgeBoveME Then
   
   lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgencias, psCtaHaber, psAreaOrig, False)
   dFecha = CDate(psFechaDocTransp)
   
   If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
      oDMov.RecuperaTipoCambioF dFecha, psOpecod, , , nTCFijo, nTCVenta, nTCComp
   End If
   
Else
   lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaDebe, psAreaDest, False)
   dFecha = CDate(psFechaDocTransp)
   If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
      oDMov.RecuperaTipoCambioF dFecha, psOpecod, , , nTCFijo, nTCVenta, nTCComp
   End If
End If
'-----------------------------------------------------------------------------------

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNroHab, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNroHab)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
If psOpecod = gOpeBoveCGConfHabAgeBoveMN Or psOpecod = gOpeBoveCGConfHabAgeBoveME Then
    lnMovItem = 0: lnMovOrden = 0
    'Guardamos el billetaje
    oDMov.GrabaMovimientoEfectivoNew lnMovNro, lnMovItem, rsBillete, psCtaDebe, "D"
    oDMov.GrabaMovimientoEfectivoNew lnMovNro, lnMovItem, rsMoneda, psCtaDebe, "D"
    lnMovOrden = 2
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(psAreaDest, 4, 2), Mid(psAreaDest, 1, 3)
    
    lnMovOrden = lnMovOrden + 1
    lnMovItem = lnMovItem + 1:

    oDMov.InsertaMovCtaPer lnMovNro, Format(lnMovItem, "#0"), psCtaHaber + lsSubCta, pnImporte * -1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(psAreaOrig, 4, 2), Mid(psAreaOrig, 1, 3)
    
Else
    lnMovItem = 0: lnMovOrden = 0
    lnMovItem = lnMovItem + 1:

    oDMov.InsertaMovCtaPer lnMovNro, Format(lnMovItem, "#0"), psCtaDebe + lsSubCta, pnImporte
    lnMovOrden = lnMovOrden + 1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(psAreaDest, 4, 2), Mid(psAreaDest, 1, 3)
    'Guardamos el billetaje
    oDMov.GrabaMovimientoEfectivoNew lnMovNro, lnMovItem, rsBillete, psCtaHaber, "H"
    oDMov.GrabaMovimientoEfectivoNew lnMovNro, lnMovItem, rsMoneda, psCtaHaber, "H"
    
    lnMovOrden = 2
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(psAreaOrig, 4, 2), Mid(psAreaOrig, 1, 3)
End If

If psNroDocSal <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocSalidaEfectivoBoveda, psNroDocSal, psFechaDocSal
End If
If psNroDocTransp <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocCompServTranspValores, psNroDocTransp, psFechaDocTransp
End If
'If psCodTransp <> "" Then
    If psAreaCodCaja = "" Then
        oDMov.InsertaMovHabilitaciones lnMovNro, psCodTransp, Mid(psAreaOrig, 1, 3), Mid(psAreaOrig, 4, 2), Mid(psAreaDest, 1, 3), Mid(psAreaDest, 4, 2), pnMontoTransp, pnImporte, Mid(psOpecod, 3, 1)
    Else
        oDMov.InsertaMovHabilitaciones lnMovNro, psCodTransp, Mid(psAreaOrig, 1, 3), Mid(psAreaOrig, 4, 2), Mid(psAreaCodCaja, 1, 3), Mid(psAreaCodCaja, 4, 2), pnMontoTransp, pnImporte, Mid(psOpecod, 3, 1)
    End If
'End If
If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovMEPerf lnMovNro, psMovNroHab, , nTCFijo, nTCVenta, nTCComp
End If

oDMov.ActualizaSaldoMovimiento psMovNroHab, "+"
oDMov.CommitTrans
lbTrans = False

GrabaHabEfectivoNew = 0
Set oDMov = Nothing
Set oContF = Nothing
Exit Function
GrabaHabEfectivoErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaHabEfectivo", Err.Description & " Error Graba Habilitación de Efectivo"
End Function


Public Function GetDatosHabilitacion(ByVal psMovNro As String) As ADODB.Recordset

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
    & " WHERE   M.CMOVNRO ='" & psMovNro & "'  AND M.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') "
    
Set rs = oConect.CargaRecordSet(sql)
Set GetDatosHabilitacion = rs
End Function
Public Function GetDatosConfHabAgencias(ByVal psMovNro As String) As ADODB.Recordset
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
    & " WHERE   M1.CMOVNRO ='" & psMovNro & "'  AND M.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') "
    
Set rs = oConect.CargaRecordSet(sql)
Set GetDatosConfHabAgencias = rs
End Function

Public Function GetHabCajaGen(ByVal psCtaCont As String, ByVal psAreaOrig As String, _
                                ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                Optional ByVal psAreaCod As String = "", Optional ByVal psAgeCod As String = "", _
                                Optional ByVal pbMuestraOrigen As Boolean = True) As ADODB.Recordset
Dim lsFiltroAge As String
If psAreaCod <> "" Then
    lsFiltroAge = " AND MH.CAREADEST ='" & psAreaCod & "' AND MH.CAGEDEST='" & psAgeCod & "' "
End If

sql = " SELECT  CONVERT(CHAR(12),CONVERT(DATETIME, SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHAHAB, " _
    & "         " & IIf(pbMuestraOrigen, " ISNULL(AG1.CAGEDESCRIPCION, A1.CAREADESCRIPCION) AS Origen, ", " ISNULL(AG.CAGEDESCRIPCION, A.CAREADESCRIPCION ) AS Destino, ") _
    & "         ISNULL(MH.NMOVIMPORTE,0 )  AS MontoHabilitado , " _
    & "         CONVERT(CHAR(25), ISNULL(C.CCONSDESCRIPCION,'')) as TipoTransp, " _
    & "         CONVERT(CHAR(50), ISNULL(P.CPERSNOMBRE,T.cDescTrans)) as Empresa, MH.nMontoTrans, M.cMovDesc, " _
    & "         A.CAREACOD , AG.CAGECOD, M.cMovNro, M.nMovNro  " _
    & " FROM    MOV M   " _
    & "         JOIN MOVHABILITACION MH ON MH.nMOVNRO=M.nMOVNRO " _
    & "         JOIN " & vsServerCom & "AREAS A1 ON A1.CAREACOD = MH.CAREAORIG " _
    & "         LEFT JOIN " & vsServerCom & "AGENCIAS AG1 ON AG1.CAGECOD= MH.CAGEORIG " _
    & "         LEFT JOIN TRANSPORTE T ON T.CTRANSPCOD = MH.CTRANSPCOD " _
    & "         LEFT JOIN " & vsServerPers & "PERSONA P ON P.CPERSCOD = T.CPERSCOD " _
    & "         LEFT JOIN " & vsServerCom & "CONSTANTE C ON C.nConsValor = T.cTpoTrans and C.nCONSCOD ='" & gCGTipoTransporte & "' " _
    & "         JOIN " & vsServerCom & "AREAS A ON A.CAREACOD =MH.cAreaDest " _
    & "         LEFT JOIN " & vsServerCom & "AGENCIAS AG ON AG.CAGECOD = MH.cAgeDest  " _
    & " WHERE   MH.nMoneda = '" & Mid(psCtaCont, 3, 1) & "' and " _
    & "         M.nMovFlag NOT IN(" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " & lsFiltroAge _
    & "         AND MH.CAREAORIG ='" & Mid(psAreaOrig, 1, 3) & "' AND MH.CAGEORIG ='" & Mid(psAreaOrig, 4, 2) & "'  " _
    & "         AND NOT EXISTS (    SELECT  M1.CMOVNRO FROM    MOV M1 " _
    & "                                     JOIN MOVREF MR ON M1.nMOVNRO =MR.nMOVNRO " _
    & "                             WHERE MR.nMOVNROREF = M.nMOVNRO AND M1.nMovFlag NOT IN(" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ")) " _
    & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND  '" & Format(pdHasta, "yyyymmdd") & "'"
    
Set rs = oConect.CargaRecordSet(sql)
Set GetHabCajaGen = rs
End Function
Public Function ExtornaHabEfectivo(ByVal pdFecsis As Date, ByVal pdFechaHab As Date, ByVal psMovNroExt As String, _
                                    ByVal pnMovNroHab As Long, ByVal psOpecod As String, ByVal psMovDesc As String, _
                                    ByVal pnImporte As Currency)
Dim oMov As DMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New DMov
On Error GoTo ExtornaHabEfectivoErr
oMov.BeginTrans
lbTrans = True
If pdFecsis = pdFechaHab Then
    oMov.InsertaMov psMovNroExt, psOpecod, psMovDesc, gMovEstContabNoContable, gMovFlagDeExtorno
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
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GrabaConfHabEfectivo(ByVal psMovNroConf As String, _
                                    ByVal rsBillete As ADODB.Recordset, _
                                    ByVal rsMoneda As ADODB.Recordset, _
                                    ByVal psOpecod As String, ByVal psMovDesc As String, _
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

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNroConf, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
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
'                    lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaDebe, psAreaDest, False)
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
'                    lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaDebe, psAreaDest, False)
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
If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovME lnMovNro, psMovNroConf
End If
oDMov.ActualizaSaldoMovimiento psMovNroConf, "+"
oDMov.CommitTrans
lbTrans = False
GrabaConfHabEfectivo = 0
Set oDMov = Nothing
Set oContF = Nothing
Exit Function
GrabaHabEfectivoErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaHabEfectivo", Err.Description & " Error Graba Habilitación de Efectivo"
End Function

Public Function GrabaConfHabEfectivoNew(ByVal psMovNroConf As String, _
                                    ByVal rsBillete As ADODB.Recordset, _
                                    ByVal rsMoneda As ADODB.Recordset, _
                                    ByVal psOpecod As String, ByVal psMovDesc As String, _
                                    ByVal psCtaDebe As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                    ByVal psAreaOrig As String, ByVal psAreaDest As String, _
                                    ByVal psMovNroHab As String) As Integer
            
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim oDMov1 As DMov
Dim oDMov2 As DMov
Dim oContF As NContFunciones
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim I As Integer
Set oDMov = New DMov
Set oDMov1 = New DMov
Set oDMov2 = New DMov
Dim lsInsertaMovCta As String
Dim lsInsertaMovCta1 As String
Dim lsInsertaMovObj As String
Dim lsInsertaMovObjEfe As String
Dim lsGeneraMovME As String
Dim lsActualizaSaldMov1 As String
Dim lsActualizaSaldMov2 As String
Dim lsActualizaSaldMov3 As String
Dim IngreBilletaje As Integer
Dim IngreMoneda As Integer

On Error GoTo GrabaHabEfectivoErr

Set oContF = New NContFunciones
GrabaConfHabEfectivoNew = 1
'---------------  JEOM ----------------------
'-----------------------------------------------------------------------------------
IngreBilletaje = 0
IngreMoneda = 0
lnMovItem = 0: lnMovOrden = 0

If Not rsBillete Is Nothing Then
    If rsBillete.State = adStateOpen Then
        If Not rsBillete.EOF And Not rsBillete.BOF Then
        I = 0
            Do While Not rsBillete.EOF
                If rsBillete!Monto > 0 Then
                   lnMovItem = lnMovItem + 1: lnMovOrden = 0
                   oDMov.InsertaMovCtaNew 15, Format(lnMovItem, "#0"), psCtaDebe + lsSubCta, rsBillete!Monto, lsInsertaMovCta
                   lnMovOrden = lnMovOrden + 1
                   oDMov.InsertaMovObjNew 15, lnMovItem, lnMovOrden, ObjDescomEfectivo, lsInsertaMovObj
                   oDMov.InsertaMovObjEfectivoNew 15, Format(lnMovItem, "#0"), Format(lnMovOrden, "#0"), rsBillete!cEfectivoCod, lsInsertaMovObjEfe
                   
                   I = I + 1
                   ReDim Preserve pMatBillete(I)
                   pMatBillete(I) = lsInsertaMovCta
                   I = I + 1
                   ReDim Preserve pMatBillete(I)
                   pMatBillete(I) = lsInsertaMovObj
                   I = I + 1
                   ReDim Preserve pMatBillete(I)
                   pMatBillete(I) = lsInsertaMovObjEfe
                   
                   IngreBilletaje = 1
                End If
                   rsBillete.MoveNext
            Loop
        End If
        'rsBillete.Close: Set rsBillete = Nothing
    End If
End If
If Not rsMoneda Is Nothing Then
    If rsMoneda.State = adStateOpen Then
        If Not rsMoneda.EOF And Not rsMoneda.BOF Then
        I = 0
            Do While Not rsMoneda.EOF
                If rsMoneda!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    oDMov.InsertaMovCtaNew 16, Format(lnMovItem, "#0"), psCtaDebe + lsSubCta, rsMoneda!Monto, lsInsertaMovCta
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObjNew 16, lnMovItem, lnMovOrden, ObjDescomEfectivo, lsInsertaMovObj
                    oDMov.InsertaMovObjEfectivoNew 16, Format(lnMovItem, "#0"), Format(lnMovOrden, "#0"), rsMoneda!cEfectivoCod, lsInsertaMovObjEfe
                    
                    I = I + 1
                    ReDim Preserve pMatMoneda(I)
                    pMatMoneda(I) = lsInsertaMovCta
                    I = I + 1
                    ReDim Preserve pMatMoneda(I)
                    pMatMoneda(I) = lsInsertaMovObj
                    I = I + 1
                    ReDim Preserve pMatMoneda(I)
                    pMatMoneda(I) = lsInsertaMovObjEfe
                    
                    IngreMoneda = 1
                End If
                rsMoneda.MoveNext
            Loop
        End If
        'rsMoneda.Close: Set rsMoneda = Nothing
    End If
End If
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaHaber, psAreaOrig, False)

lnMovItem = lnMovItem + 1: lnMovOrden = 0
lnMovOrden = lnMovOrden + 1

oDMov.InsertaMovCtaNew 17, Format(lnMovItem, "#0"), psCtaHaber + lsSubCta, pnImporte * -1, lsInsertaMovCta1
ReDim pMatMovCta(1)
pMatMovCta(1) = lsInsertaMovCta1

If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovMENew 18, psMovNroConf, , lsGeneraMovME
End If
ReDim pMatMovME(1)
pMatMovME(1) = lsGeneraMovME


oDMov.ActualizaSaldoMovimientoNew psMovNroConf, "+", lsActualizaSaldMov1, lsActualizaSaldMov2, lsActualizaSaldMov3

ReDim pMatActSaldMov1(1)
pMatActSaldMov1(1) = lsActualizaSaldMov1

ReDim pMatActSaldMov2(1)
pMatActSaldMov2(1) = lsActualizaSaldMov2

ReDim pMatActSaldMov3(1)
pMatActSaldMov3(1) = lsActualizaSaldMov3

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

oDMov.BeginTrans
    lbTrans = True
    
    oDMov.InsertaMov psMovNroConf, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
    lnMovNro = oDMov.GetnMovNro(psMovNroConf)
    
    I = 1
    If IngreBilletaje = 1 Then
        For I = 1 To UBound(pMatBillete) - 1
            pMatBillete(I) = Replace(pMatBillete(I), "15", Trim(Str(lnMovNro)))
            oDMov.EjecutarNew (pMatBillete(I))
        Next I
    rsBillete.Close: Set rsBillete = Nothing
    End If
    
    
    I = 1
    If IngreMoneda = 1 Then
        For I = 1 To UBound(pMatMoneda) - 1
            pMatMoneda(I) = Replace(pMatMoneda(I), "16", Trim(Str(lnMovNro)))
            oDMov.EjecutarNew (pMatMoneda(I))
        Next I
    rsMoneda.Close: Set rsMoneda = Nothing
    End If
    
    
    oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
    

    
    'insertamos la cuenta del haber
    
    lnMovItem = lnMovItem + 1: lnMovOrden = 0
    lnMovOrden = lnMovOrden + 1
    pMatMovCta(1) = Replace(pMatMovCta(1), "17", Trim(Str(lnMovNro)))
    oDMov.EjecutarNew (pMatMovCta(1))
    lnMovOrden = lnMovOrden + 1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
    oDMov.InsertaMovObjAgenciaArea lnMovNro, Format(lnMovItem, "#0"), Format(lnMovOrden, "#0"), Mid(psAreaOrig, 4, 2), Mid(psAreaOrig, 1, 3)
    
    oDMov.InsertaMovRef lnMovNro, psMovNroHab
    
    If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
        pMatMovME(1) = Replace(pMatMovME(1), "18", Trim(Str(lnMovNro)))
        oDMov.EjecutarNew pMatMovME(1)
    End If
    
    'oDMov.ActualizaSaldoMovimiento psMovNroConf, "+"
    oDMov.EjecutarNew pMatActSaldMov1(1)
    oDMov.EjecutarNew pMatActSaldMov2(1)
    oDMov.EjecutarNew pMatActSaldMov3(1)


oDMov.CommitTrans
lbTrans = False
GrabaConfHabEfectivoNew = 0
Set oDMov = Nothing
Set oContF = Nothing
Exit Function
GrabaHabEfectivoErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaHabEfectivo", Err.Description & " Error Graba Habilitación de Efectivo"
End Function
Public Function GetDatosConfHabilitacion(ByVal psOpecod As String, ByVal psAreaOrig As String, _
                                        ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                        Optional ByVal psAreaCod As String = "", Optional ByVal psAgeCod As String = "") As ADODB.Recordset
Dim lsFiltroAge As String

If psAreaCod <> "" Then
    lsFiltroAge = " AND MOAA.CAREACOD ='" & psAreaCod & "' AND MOAA.CAGECOD='" & psAgeCod & "' "
End If

sql = "SELECT  CONVERT(CHAR(12),CONVERT(DATETIME, SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHAHAB, " _
    & "        ISNULL(AG1.CAGEDESCRIPCION, A1.CAREADESCRIPCION ) AS Destino, " _
    & "        SUM( ISNULL(ME.NMOVMEIMPORTE , MC.NMOVIMPORTE ) )  AS MontoHabilitado , " _
    & "            '' as TipoTransp, '' as Empresa, " _
    & "            0 nMontoTrans, M.cMovDesc, '' CAREACOD, '' CAGECOD, M.cMovNro, M.nMovNro " _
    & "    from mov m " _
    & "    JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO " _
    & "             LEFT JOIN MOVME ME ON ME.nMOVNRO =MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "             JOIN MovObjEfectivo mef ON mef.nMovNro = mc.nMovNro and mef.nMovItem = mc.nMovItem " _
    & "             JOIN MovObjAreaAgencia mh ON m.nMovNro = mh.nMovNro " _
    & "             JOIN AREAS A1 ON A1.CAREACOD = MH.CAREACod " _
    & "             LEFT JOIN AGENCIAS AG1 ON AG1.CAGECOD= MH.CAGECod "
    
sql = sql + " WHERE   M.COPECOD = '" & psOpecod & "' and mc.nMovImporte > 0 " _
    & "         AND M.nMovFlag NOT  IN(" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " & lsFiltroAge _
    & "         AND MH.CAREACod ='" & Mid(psAreaOrig, 1, 3) & "' AND MH.CAGECod ='" & Mid(psAreaOrig, 4, 2) & "' " _
    & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND  '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "Group By CONVERT(CHAR(12),CONVERT(DATETIME, SUBSTRING(M.CMOVNRO,1,8)),103), " _
    & " ISNULL(AG1.CAGEDESCRIPCION, A1.CAREADESCRIPCION ), m.cMovDesc , m.cMovNro, m.nMovNro "
Set rs = oConect.CargaRecordSet(sql)
Set GetDatosConfHabilitacion = rs
End Function
Public Function GrabaCompraVentaMEIF(ByVal psMovNro As String, _
                                ByVal psOpecod As String, ByVal psMovDesc As String, _
                                ByVal psCtaDebe As String, ByVal psCtaHaber As String, _
                                ByVal psCtaIfOrig As String, ByVal psCtaIfDest As String, _
                                ByVal psTpoDocOrig As TpoDoc, ByVal psNroDocOrig As String, _
                                ByVal psTpoDocDest As TpoDoc, ByVal psNroDocDest As String, _
                                ByVal pnImporte As Currency, ByVal pnTCBanco As Currency, ByVal pdFecsis As Date, Optional ByVal pnTipoCambio As Currency = 0#)

Dim oMov As DMov
Dim lnItem As Integer
Dim lnOrden As Integer
Dim lbTrans As Boolean
Dim lnMovNro As Long

'ALPA 20140226*******************

Dim objTC As NCompraVenta
Set objTC = New NCompraVenta
Dim oRsTCambio As ADODB.Recordset
Set oRsTCambio = New ADODB.Recordset

Dim lnValVent As Currency
Dim lnValComp As Currency
Dim lnValFijo As Currency
Dim lnValFijoDia As Currency
Dim lnValVentEsp As Currency
Dim lnValCompEsp As Currency
Dim lnValPond As Currency
Dim lnValPondVenta As Currency
Dim lnValPondREU As Currency
Dim lnValSBSDia As Currency
Dim lnValCompTr As Currency
Dim lnValVentTr As Currency
'********************************
On Error GoTo GrabaCompraVentaMEIFErr

'ALPA 20140226*******************
Set oRsTCambio = objTC.ObtenerTipoCambioDiario(CDate(Mid(psMovNro, 1, 4) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 7, 2)))
If Not (oRsTCambio.BOF Or oRsTCambio.EOF) Then
    Do While Not oRsTCambio.EOF
        lnValVent = oRsTCambio!nValVent
        lnValComp = oRsTCambio!nValComp
        lnValFijo = oRsTCambio!nValFijo
        lnValFijoDia = oRsTCambio!nValFijoDia
        lnValVentEsp = oRsTCambio!nValVentEsp
        lnValCompEsp = oRsTCambio!nValCompEsp
        lnValPond = oRsTCambio!nValPond
        lnValPondVenta = oRsTCambio!nValPondVenta
        lnValPondREU = oRsTCambio!nValPondREU
        lnValSBSDia = oRsTCambio!nValSBSDia
        lnValCompTr = oRsTCambio!nValCompTr
        lnValVentTr = oRsTCambio!nValVentTr
        oRsTCambio.MoveNext
    Loop
End If
'********************************

Set oMov = New DMov
oMov.BeginTrans
lbTrans = True
oMov.InsertaMov psMovNro, psOpecod, psMovDesc
lnMovNro = oMov.GetnMovNro(psMovNro)

oMov.InsertaMovCont lnMovNro, pnImporte, 0, "0"
lnItem = 0
lnItem = lnItem + 1: lnOrden = 0
oMov.InsertaMovCta lnMovNro, lnItem, psCtaDebe, pnImporte
lnOrden = lnOrden + 1
oMov.InsertaMovObj lnMovNro, lnItem, lnOrden, ObjEntidadesFinancieras
oMov.InsertaMovObjIF lnMovNro, lnItem, lnOrden, Mid(psCtaIfDest, 4, 13), Mid(psCtaIfDest, 1, 2), Mid(psCtaIfDest, 18, 10)

'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
'oMov.InsertaMovCtaIF lnMovNro, Mid(psCtaIfOrig, 4, 13), Mid(psCtaIfOrig, 1, 2), Mid(psCtaIfOrig, 18, 10), CGCtaIFConCapital, pnImporte
'oMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
'                        Mid(psCtaIfOrig, 4, 13), Mid(psCtaIfOrig, 1, 2), Mid(psCtaIfOrig, 18, 10)
                        
lnItem = lnItem + 1: lnOrden = 0
oMov.InsertaMovCta lnMovNro, lnItem, psCtaHaber, pnImporte * -1
lnOrden = lnOrden + 1
oMov.InsertaMovObj lnMovNro, lnItem, lnOrden, ObjEntidadesFinancieras
oMov.InsertaMovObjIF lnMovNro, lnItem, lnOrden, Mid(psCtaIfOrig, 4, 13), Mid(psCtaIfOrig, 1, 2), Mid(psCtaIfOrig, 18, 10)

'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
'oMov.InsertaMovCtaIF lnMovNro, Mid(psCtaIfDest, 4, 13), Mid(psCtaIfDest, 1, 2), Mid(psCtaIfDest, 18, 10), CGCtaIFConCapital, pnImporte * -1

'Select Case psOpeCod
'    Case gOpeMECompraAInst, gOpeMEVentaAInst
'         oMov.GeneraMovME lnMovNro, psMovNro, pnTCBanco
'End Select

oMov.GeneraMovME lnMovNro, psMovNro, pnTCBanco, , , , True, lnValVent, lnValComp, lnValFijo, lnValFijoDia, lnValVentEsp, lnValCompEsp, lnValPond, lnValPondVenta, lnValPondREU, lnValSBSDia, lnValCompTr, lnValVentTr

If psNroDocOrig <> "" Then
    oMov.InsertaMovDoc lnMovNro, psTpoDocOrig, psNroDocOrig, Format(pdFecsis, gsFormatoFecha)
End If
If psNroDocDest <> "" Then
    oMov.InsertaMovDoc lnMovNro, psTpoDocDest, psNroDocDest, Format(pdFecsis, gsFormatoFecha)
End If

oMov.ActualizaSaldoMovimiento psMovNro, "+"

oMov.CommitTrans
lbTrans = False
Set oMov = Nothing
Exit Function
GrabaCompraVentaMEIFErr:
   lsMsgErr = Err.Description
        If lbTrans Then
            oMov.RollbackTrans
        End If
        lbTrans = False
        Err.Raise 50001, "nCajaGeneral: GrabaCompraVentaMEIF", lsMsgErr
End Function
Public Function GrabaCompraVentaEfectivo(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
                                    ByVal rsBillIngreso As ADODB.Recordset, ByVal rsMonIng As ADODB.Recordset, _
                                    ByVal rsBillSale As ADODB.Recordset, ByVal rsMonSale As ADODB.Recordset, _
                                    ByVal psCtaDebe As String, ByVal psCtaHaber As String, _
                                    ByVal pnTCBanco As Currency, pnTipCambio As Currency) As Integer
            
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov     As DMov
Dim oContF    As NContFunciones
Dim lsMovNro  As String
Dim lnMovNro  As Long
Dim lsSubCta  As String
Dim lbTrans   As Boolean
Dim lnImporte As Currency, lnTotalD As Currency, lnTotalS As Currency

Set oDMov = New DMov
On Error GoTo GrabaCompraVentaEfectivoErr

Set oContF = New NContFunciones

GrabaCompraVentaEfectivo = 1

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, 0, "0", "0"

lnTotalD = 0: lnTotalS = 0
'guardamos la cuenta en el haber
lnMovItem = 0: lnMovOrden = 0
'guardamos la cuenta en el debe
If Not rsBillIngreso Is Nothing Then
    If rsBillIngreso.State = adStateOpen Then
        If Not rsBillIngreso.EOF And Not rsBillIngreso.BOF Then
            Do While Not rsBillIngreso.EOF
                If rsBillIngreso!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    If psOpecod = gOpeMECompraEfect Then
                       lnImporte = Round(rsBillIngreso!Monto * pnTipCambio, 2)
                       lnTotalD = lnTotalD + lnImporte
                       oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe, lnImporte
                       oDMov.InsertaMovMe lnMovNro, Format(lnMovItem, "#0"), rsBillIngreso!Monto
                    Else
                       oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe, rsBillIngreso!Monto
                       lnTotalS = lnTotalS + rsBillIngreso!Monto
                    End If
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
                    If psOpecod = gOpeMECompraEfect Then
                       lnImporte = Round(rsMonIng!Monto * pnTipCambio, 2)
                       lnTotalD = lnTotalD + lnImporte
                       oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, lnImporte
                       oDMov.InsertaMovMe lnMovNro, lnMovItem, rsMonIng!Monto
                    Else
                       oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, rsMonIng!Monto
                       lnTotalS = lnTotalS + rsMonIng!Monto
                    End If
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
                    If psOpecod = gOpeMEVentaEfec Then
                       lnImporte = Round(rsBillSale!Monto * pnTipCambio, 2)
                       lnTotalD = lnTotalD + lnImporte
                       oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, lnImporte * -1
                       oDMov.InsertaMovMe lnMovNro, lnMovItem, rsBillSale!Monto * -1
                    Else
                       oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, rsBillSale!Monto * -1
                       lnTotalS = lnTotalS + rsBillSale!Monto
                    End If
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
If Not rsMonSale Is Nothing Then
    If rsMonSale.State = adStateOpen Then
        If Not rsMonSale.EOF And Not rsMonSale.BOF Then
            Do While Not rsMonSale.EOF
                If rsMonSale!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    If psOpecod = gOpeMEVentaEfec Then
                       lnImporte = Round(rsMonSale!Monto * pnTipCambio, 2)
                       lnTotalD = lnTotalD + lnImporte
                       oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, lnImporte * -1
                       oDMov.InsertaMovMe lnMovNro, lnMovItem, rsMonSale!Monto * -1
                    Else
                       oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, rsMonSale!Monto * -1
                       lnTotalS = lnTotalS + rsMonSale!Monto
                    End If
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, rsMonSale!cEfectivoCod
                End If
                rsMonSale.MoveNext
            Loop
        End If
        rsMonSale.Close: Set rsMonSale = Nothing
    End If
End If

If psOpecod = gOpeMEVentaEfec Or psOpecod = gOpeMECompraEfect Then
   If lnTotalD <> lnTotalS Then
      Dim sConvMES    As String
      Dim sConvMED    As String
      Dim nDif        As Currency
      Dim oCons As New NConstSistemas
      sConvMED = oCons.LeeConstSistema(gConstSistCtaConversionMEDol)
      sConvMES = oCons.LeeConstSistema(gConstSistCtaConversiónMESoles)
      Set oCons = Nothing
      lnMovItem = lnMovItem + 1
      If psOpecod = gOpeMEVentaEfec Then
         nDif = lnTotalD - lnTotalS
      Else
         nDif = lnTotalS - lnTotalD
      End If
      If nDif > 0 Then 'Perdida
         oDMov.InsertaMovCta lnMovNro, lnMovItem, sConvMED, nDif
      Else                             'Ganancia
         oDMov.InsertaMovCta lnMovNro, lnMovItem, sConvMES, nDif
      End If
   End If
End If
If psOpecod = OpeCGOtrosOpeEfecCambme Then
   oDMov.GeneraMovME lnMovNro, psMovNro, IIf(pnTCBanco = 1, 0, pnTCBanco)
Else
   oDMov.InsertaMovTpoCambio lnMovNro, pnTCBanco
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"
oDMov.CommitTrans
GrabaCompraVentaEfectivo = 0
lbTrans = False
Set oDMov = Nothing
Set oContF = Nothing
Exit Function
GrabaCompraVentaEfectivoErr:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
    End If
    Err.Raise vbObjectError + 100, "GrabaCompraVentaEfectivo", lsMsgErr
End Function
Public Function GetMovCompraVentaME(ByVal psOpecod As String, ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset
Dim lsFiltroAge As String

sql = " SELECT  CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) as Fecha ," _
    & "         D.CPERSNOMBRE, D.CTAIFDEST, H.cPersNombre , H.CTAIFORIG, D.IMPORTE, M.nMovNro, M.cMovDesc, M.cMovNro " _
    & " FROM    MOV M " _
    & "         JOIN (  SELECT  M.CMOVNRO , M.nMOVNRO  , P.cPersNombre , C1.cCtaIFDesc ,  C.cCtaIFDesc AS CTAIFORIG , " _
    & "                         ISNULL(Me.NMOVMEIMPORTE, MC.NMOVIMPORTE) As IMPORTE " _
    & "                 FROM    MOV M  JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO  " _
    & "                         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "                         JOIN MOVOBJIF MO ON MO.nMOVNRO = MC.nMOVNRO AND MO.nMOVITEM = MC.nMOVITEM " _
    & "                         JOIN CTAIF C ON C.CPERSCOD = MO.CPERSCOD AND MO.CCTAIFCOD = C.CCTAIFCOD  AND C.CIFTPO = MO.CIFTPO   " _
    & "                         JOIN CTAIF C1 ON C1.CPERSCOD = MO.CPERSCOD AND C1.CIFTPO = MO.CIFTPO AND MO.CCTAIFCOD LIKE C1.cCtaIFCod+'%' AND LEN(C1.cCtaIFCod) < LEN (MO.CCTAIFCOD) " _
    & "                         JOIN " & vsServerPers & "PERSONA P ON P.CPERSCOD = C.CPERSCOD " _
    & "                 WHERE   MC.NMOVIMPORTE <0  AND M.COPECOD ='" & psOpecod & "'  ) AS H " _
    & "         ON H.nMOVNRO = M.nMOVNRO " _
    & "         JOIN ( SELECT   M.CMOVNRO, M.nMOVNRO ,  P.cPersNombre,  C1.cCtaIFDesc  ,  C.cCtaIFDesc  AS CTAIFDEST, " _
    & "                         ISNULL(Me.NMOVMEIMPORTE, MC.NMOVIMPORTE) As IMPORTE " _
    & "                 FROM    MOV M  JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO  " _
    & "                         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "                         JOIN MOVOBJIF MO ON MO.nMOVNRO = MC.nMOVNRO AND MO.nMOVITEM = MC.nMOVITEM " _
    & "                         JOIN CTAIF C ON C.CPERSCOD = MO.CPERSCOD AND MO.CCTAIFCOD = C.CCTAIFCOD AND C.CIFTPO = MO.CIFTPO " _
    & "                         JOIN CTAIF C1 ON C1.CPERSCOD = MO.CPERSCOD AND C1.CIFTPO = MO.CIFTPO AND MO.CCTAIFCOD LIKE C1.cCtaIFCod+'%' AND LEN(C1.cCtaIFCod) < LEN (MO.CCTAIFCOD) " _
    & "                         JOIN " & vsServerPers & "PERSONA P ON P.CPERSCOD = C.CPERSCOD " _
    & "                 WHERE   MC.NMOVIMPORTE >0 AND M.COPECOD ='" & psOpecod & "' ) AS D " _
    & "         ON D.nMOVNRO = M.nMOVNRO " _
    & " WHERE   M.COPECOD ='" & psOpecod & "' AND M.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "')" _
    & "         AND SUBSTRING(M.cMovnro,1,8 ) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "'  ORDER BY M.CMOVNRO "


Set rs = oConect.CargaRecordSet(sql)
Set GetMovCompraVentaME = rs

End Function
Public Function GrabaExtornoMov(ByVal pdFecsis As Date, ByVal pdFechaMov As Date, ByVal psMovNroExt As String, _
                                  ByVal pnMovNroAnt As Long, ByVal psOpecod As String, ByVal psMovDesc As String, _
                                  ByVal pnImporte As Currency, ByVal psMovActualiza As String, ByVal lbEliminaMov As Boolean, _
                                  Optional psMovNroAnt As String = "", Optional psAgeCodRef As String = "", Optional psDocNro As String = "", _
                                  Optional psObjetoCod As String = "", Optional pbBitCentral As Boolean = True, Optional pdFecReg As Date, Optional pnNumTran As Long = 0, Optional pnImporteIntDev As Currency = 0)

On Error GoTo ErrorGrabaExtornoMov
Dim oMov As DMov
Dim lbTrans As Boolean
Dim oChq As New NDocRec

'PASIERS1242014
Dim oCapta As NCapMovimientos
Dim oDCapta As DCapMantenimiento
Dim rsMovNeg As ADODB.Recordset
Dim lsMovNroNegocio As String
'end PASI

Set oMov = New DMov
oMov.BeginTrans
lbTrans = True
   Select Case psOpecod
    'Operaciones con Ventanilla
    Case gOpeCGRVentanaIngresoMNExt, gOpeCGRVentanaEgresoMNExt, _
         gOpeCGRVentanaIngresoMEExt, gOpeCGRVentanaEgresoMEExt
         oMov.ExtornaMovimiento psMovNroExt, pnMovNroAnt, psOpecod, psMovDesc, lbEliminaMov, psMovActualiza
         If Not pbBitCentral Then
            oMov.ExtornaMovimientoRefNegocio pnNumTran, psMovNroAnt, psAgeCodRef
         End If
         
    'Extornos de Depositos/Retiros Diversos de Bancos
    Case gOpeCGExtBcoDepDiv, gOpeCGExtBcoDepDivME, _
         gOpeCGExtBcoRetDiv, gOpeCGExtBcoRetDivME, _
         gOpeCGOpeBancosDepositoXActivacionSegTarjetaMNExt, _
         gOpeCGOpeBancosRetPagoSegTarjMNExt, gOpeCGOpeBancosDepComSegTarjMNExt, _
         "421018"
         
         'FRHU 20140528 ERS068-2014 gOpeCGOpeBancosDepositoXActivacionSegTarjetaMNExt
         'JUEZ 20140711 gOpeCGOpeBancosRetPagoSegTarjMNExt, gOpeCGOpeBancosDepComSegTarjMNExt
         'APRI 20180130 ERS028-2017 ADD 421018"
         
         oMov.ExtornaMovimiento psMovNroExt, pnMovNroAnt, psOpecod, psMovDesc, lbEliminaMov, psMovActualiza
        
        'Para el Caso de Notas de Abonota/Cargo..
        'el extorno se debe hacer en el Sicmac Negocio
        
        'Para las Operaciones del F5 volver a Null TransRef
         If Not pbBitCentral Then
            oMov.ExtornaMovimientoRefNegocio pnNumTran, psMovNroAnt, psAgeCodRef
         End If

    'Extornos de Ope de Cheques
    Case gOpeCGExtBcoDepCheques, gOpeCGExtBcoDepChequesME
         oMov.ExtornaMovimiento psMovNroExt, pnMovNroAnt, psOpecod, psMovDesc, lbEliminaMov, psMovActualiza
         oMov.ExtornaDepositoCheque pnMovNroAnt
         
    Case gOpeCGExtBcoRegCheques, gOpeCGExtBcoRegChequesME
         oChq.GrabaExtornoCheque pnMovNroAnt, psMovNroExt, psOpecod, "", psMovDesc, psDocNro, Mid(psObjetoCod, 4, 13), Left(psObjetoCod, 2), pnImporte, , , , lbEliminaMov, Mid(psObjetoCod, 18)
         
    Case gOpeCGExtBcoRecepChqRegAgencias, gOpeCGExtBcoRecepChqRegAgenciasME
         oChq.GrabaExtornoCheque pnMovNroAnt, psMovNroExt, psOpecod, "", psMovDesc, psDocNro, Mid(psObjetoCod, 4, 13), Left(psObjetoCod, 2), pnImporte, , , , lbEliminaMov, Mid(psObjetoCod, 18)
         If Not pbBitCentral Then
            oMov.ExtornaConfirmacionChequeAgencia psAgeCodRef, psDocNro, pdFecReg, psObjetoCod
         End If
          
    Case gOpeCGExtBcoApertCta, gOpeCGExtBcoApertCtaME, _
         gOpeCGExtBcoConfApert, gOpeCGExtBcoConfApertME, _
         gOpeCGExtBcoCancelaCtas, gOpeCGExtBcoCancelaCtasME, _
         gOpeCGExtBcoCapitalizaIntDPF, gOpeCGExtBcoCapitalizaIntDPFME
         oMov.ExtornaMovimiento psMovNroExt, pnMovNroAnt, psOpecod, psMovDesc, lbEliminaMov, psMovActualiza
         Select Case psOpecod
            Case gOpeCGExtBcoConfApert, gOpeCGExtBcoConfApertME
               oMov.ActualizaCtaIF Mid(psObjetoCod, 4, 13), Left(psObjetoCod, 2), Mid(psObjetoCod, 18), , , , , , , gEstadoCtaIFRegistrada
               oMov.EliminaCuentaIFFiltroPersona Mid(psObjetoCod, 4, 13), Left(psObjetoCod, 2), Mid(psObjetoCod, 18)
            
            Case gOpeCGExtBcoApertCta, gOpeCGExtBcoApertCtaME
               oMov.ActualizaCtaIF Mid(psObjetoCod, 4, 13), Left(psObjetoCod, 2), Mid(psObjetoCod, 18), , , , , , , gEstadoCtaIFAnulada
            Case gOpeCGExtBcoCancelaCtas, gOpeCGExtBcoCancelaCtasME, _
                 gOpeCGExtBcoCapitalizaIntDPF, gOpeCGExtBcoCapitalizaIntDPFME
              'Falta deducir Ult. Capitalizacion
               oMov.ActualizaCtaIF Mid(psObjetoCod, 4, 13), Left(psObjetoCod, 2), Mid(psObjetoCod, 18), , , , , , , gEstadoCtaIFActiva, pnImporteIntDev
         End Select
         
    'PF Int Devengados
    Case gOpeCGExtBcoIntDevengPF, gOpeCGExtBcoIntDevengPFME
         oMov.ExtornaMovimiento psMovNroExt, pnMovNroAnt, psOpecod, psMovDesc, lbEliminaMov, psMovActualiza
         oMov.ExtornaProvisionPlazoFijo psMovNroAnt, psOpecod, pnImporte, psObjetoCod

    'Extornos Adeudados
    Case gOpeCGAdeudaExtRegistroMN, gOpeCGAdeudaExtRegistroME, _
         gOpeCGAdeudaExtConfRegiMN, gOpeCGAdeudaExtConfRegiME
         oMov.ExtornaMovimiento psMovNroExt, pnMovNroAnt, psOpecod, psMovDesc, lbEliminaMov, psMovActualiza
         If psOpecod = gOpeCGAdeudaExtConfRegiMN Or psOpecod = gOpeCGAdeudaExtConfRegiME Then
            oMov.ActualizaCtaIF Mid(psObjetoCod, 4, 13), Left(psObjetoCod, 2), Mid(psObjetoCod, 18), , , , , , , gEstadoCtaIFRegistrada
         Else
            oMov.ActualizaCtaIF Mid(psObjetoCod, 4, 13), Left(psObjetoCod, 2), Mid(psObjetoCod, 18), , , , , , , gEstadoCtaIFAnulada
         End If
    'Extorno de Provisiones y Pago de Cuotas de Adeudados
    Case gOpeCGAdeudaExtProvisiónMN, gOpeCGAdeudaExtProvisiónME, _
         gOpeCGAdeudaExtPagoCuotaMN, gOpeCGAdeudaExtPagoCuotaME
         oMov.ExtornaMovimiento psMovNroExt, pnMovNroAnt, psOpecod, psMovDesc, lbEliminaMov, psMovActualiza
         oMov.ExtornaPagoCuotaAdeudados psMovNroAnt, psOpecod, pnImporte, pdFecReg, pnMovNroAnt, ""
    
    Case gOpeCGExtProvPagoAbono, gOpeCGExtProvPagoAbonoME 'PASIERS1242014
         Set oCapta = New NCapMovimientos
         Set oDCapta = New DCapMantenimiento
         Set rsMovNeg = New ADODB.Recordset
         Set rsMovNeg = oDCapta.GetMovNegxExtornoPagoProveedor("200207", pnMovNroAnt)
         oMov.ExtornaMovimiento psMovNroExt, pnMovNroAnt, psOpecod, psMovDesc, lbEliminaMov, psMovActualiza
         lsMovNroNegocio = oMov.GeneraMovNro(, , , psMovNroExt)
         oCapta.CapExtornoAbonoAho rsMovNeg!nMovNro, "230207", rsMovNeg!cCtaCod, lsMovNroNegocio, psMovDesc, Format$(rsMovNeg!nMonto, "#,##0.00"), , , , gsNomAge, , Format$(rsMovNeg!ITFCargo, "#,##0.00"), Trim(rsMovNeg!ITFOperacion), Trim(rsMovNeg!ITFConcepto), True, oMov.GetConexion, False
         Set oCapta = Nothing
         Set oDCapta = Nothing
         Set rsMovNeg = Nothing
         'END PASI
    Case gOpeCGExtProvDevSUNAT 'PASIERS1242014
        Dim lsMovNroExt2 As String
        Dim lnMovNro2 As Long
        Dim rsMov As ADODB.Recordset
        
        
         Set oCapta = New NCapMovimientos
         Set oDCapta = New DCapMantenimiento
         Set rsMovNeg = New ADODB.Recordset
         Set rsMov = New ADODB.Recordset
         
         
         Set rsMov = oDCapta.ObtieneMovIniDevSUNAT(pnMovNroAnt)
         
         Do While Not rsMov.EOF
            lsMovNroExt2 = oMov.GeneraMovNro(, , , IIf(Len(lsMovNroExt2) = 0, psMovNroExt, lsMovNroExt2))
            oMov.ExtornaMovimiento lsMovNroExt2, rsMov!nMovNro, psOpecod, psMovDesc, lbEliminaMov, psMovActualiza
            rsMov.MoveNext
         Loop
         Set rsMovNeg = oDCapta.GetMovNegxExtornoPagoProveedor("200207", pnMovNroAnt)
         oMov.ExtornaMovimiento psMovNroExt, pnMovNroAnt, psOpecod, psMovDesc, lbEliminaMov, psMovActualiza
         If Not rsMovNeg.EOF Then
            lsMovNroNegocio = oMov.GeneraMovNro(, , , psMovNroExt)
            oCapta.CapExtornoAbonoAho rsMovNeg!nMovNro, "230207", rsMovNeg!cCtaCod, lsMovNroNegocio, psMovDesc, Format$(rsMovNeg!nMonto, "#,##0.00"), , , , gsNomAge, , Format$(rsMovNeg!ITFCargo, "#,##0.00"), Trim(rsMovNeg!ITFOperacion), Trim(rsMovNeg!ITFConcepto), True, oMov.GetConexion, False
         End If
         
         Set oCapta = Nothing
         Set oDCapta = Nothing
         Set rsMovNeg = Nothing
    
    Case Else
         oMov.ExtornaMovimiento psMovNroExt, pnMovNroAnt, psOpecod, psMovDesc, lbEliminaMov, psMovActualiza
    End Select
oMov.CommitTrans
lbTrans = False
Set oMov = Nothing
Set oChq = Nothing
Exit Function
ErrorGrabaExtornoMov:
    lsMsgErr = Err.Description
    If lbTrans Then oMov.RollbackTrans
    Err.Raise Err.Number, "GrabaExtornoMov", lsMsgErr
End Function


Public Function GetMovCompraVentaEfectivo(ByVal psOpecod As String, ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset
Dim lsFiltroAge As String

sql = " SELECT  CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) as Fecha ," _
    & "         M.CMOVDESC, D.MEIMPORTE AS DOLARES, S.MNIMPORTE , MTC.nMovTpoCambio, S.cMovNro, S.nMOVNRO " _
    & " FROM    MOV M " _
    & "         JOIN (  SELECT  M.CMOVNRO, M.nMOVNRO , SUM(ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE))) AS MEIMPORTE " _
    & "                 FROM    MOV M JOIN MOVCTA MC  ON MC.nMOVNRO =M.nMOVNRO " _
    & "                         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "                 WHERE   MC.CCTACONTCOD LIKE '__" & gMonedaExtranjera & "%'  AND M.COPECOD ='" & psOpecod & "'  " _
    & "                 GROUP  BY  M.cMOVNRO, M.nMOVNRO   ) AS D " _
    & "         ON D.nMOVNRO = M.nMOVNRO " _
    & "         JOIN (  SELECT  M.CMOVNRO, M.nMOVNRO, SUM(ISNULL(ABS(MC.NMOVIMPORTE),0)) AS MNIMPORTE " _
    & "                 FROM    MOV M JOIN MOVCTA MC  ON MC.nMOVNRO =M.nMOVNRO " _
    & "                         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "                 WHERE   MC.CCTACONTCOD LIKE '__" & gMonedaNacional & "%' AND M.COPECOD ='" & psOpecod & "'  " _
    & "                 GROUP  BY  M.CMOVNRO, M.nMOVNRO   ) AS S " _
    & "         ON S.nMOVNRO = M.nMOVNRO " _
    & "         JOIN MOVTPOCAMBIO  MTC ON MTC.nMOVNRO = M.nMOVNRO " _
    & " WHERE   M.COPECOD ='" & psOpecod & "' AND M.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "')" _
    & "         AND SUBSTRING(S.cMovnro,1,8 ) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' ORDER BY M.CMOVNRO "

Set rs = oConect.CargaRecordSet(sql)
Set GetMovCompraVentaEfectivo = rs

End Function
Public Function GrabaMovimientosCtasIF(ByVal psMovNro As String, _
                                        ByVal psOpecod As String, ByVal psMovDesc As String, _
                                        ByVal psCtaDebe As String, ByVal psCtaHaber As String, _
                                        ByVal psCtaIfOrig As String, ByVal psCtaIfDest As String, _
                                        ByVal psTpoDocOrig As TpoDoc, ByVal psNroDocOrig As String, _
                                        ByVal psTpoDocDest As TpoDoc, ByVal psNroDocDest As String, _
                                        ByVal pnImporte As Currency, ByVal pnTCBanco As Currency, ByVal pdFecsis As Date)

Dim oMov As DMov
Dim lnItem As Integer
Dim lnOrden As Integer
Dim lbTrans As Boolean
Dim lnMovNro As Long
On Error GoTo GrabaMovimientosCtasIFErr

Set oMov = New DMov

GrabaMovimientosCtasIF = 1

oMov.BeginTrans
lbTrans = True
oMov.InsertaMov psMovNro, psOpecod, psMovDesc
lnMovNro = oMov.GetnMovNro(psMovNro)

oMov.InsertaMovCont lnMovNro, pnImporte, 0, "0"
lnItem = 0
lnItem = lnItem + 1: lnOrden = 0
oMov.InsertaMovCta lnMovNro, lnItem, psCtaDebe, pnImporte
lnOrden = lnOrden + 1
oMov.InsertaMovObj lnMovNro, lnItem, lnOrden, ObjEntidadesFinancieras
oMov.InsertaMovObjIF lnMovNro, lnItem, lnOrden, Mid(psCtaIfDest, 4, 13), Mid(psCtaIfDest, 1, 2), Mid(psCtaIfDest, 18, 10)

'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
'oMov.InsertaMovCtaIF lnMovNro, Mid(psCtaIfOrig, 4, 13), Mid(psCtaIfOrig, 1, 2), Mid(psCtaIfOrig, 18, 10), CGCtaIFConCapital, pnImporte
'oMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                        Mid(psCtaIfDest, 4, 13), Mid(psCtaIfDest, 1, 2), Mid(psCtaIfDest, 18, 10)

lnItem = lnItem + 1: lnOrden = 0
oMov.InsertaMovCta lnMovNro, lnItem, psCtaHaber, pnImporte * -1
lnOrden = lnOrden + 1
oMov.InsertaMovObj lnMovNro, lnItem, lnOrden, ObjEntidadesFinancieras
oMov.InsertaMovObjIF lnMovNro, lnItem, lnOrden, Mid(psCtaIfOrig, 4, 13), Mid(psCtaIfOrig, 1, 2), Mid(psCtaIfOrig, 18, 10)

'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
'oMov.InsertaMovCtaIF lnMovNro, Mid(psCtaIfDest, 4, 13), Mid(psCtaIfDest, 1, 2), Mid(psCtaIfDest, 18, 10), CGCtaIFConCapital, pnImporte * -1
'oMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                        Mid(psCtaIfDest, 4, 13), Mid(psCtaIfDest, 1, 2), Mid(psCtaIfDest, 18, 10)

'If pnTCBanco <> 0 Then
'    oMov.InsertaMovTpoCambio lnMovNro, pnTCBanco
'End If
If psNroDocOrig <> "" Then
   oMov.InsertaMovDoc lnMovNro, psTpoDocOrig, psNroDocOrig, Format(pdFecsis, gsFormatoFecha)
End If
If psNroDocDest <> "" Then
   oMov.InsertaMovDoc lnMovNro, psTpoDocDest, psNroDocDest, Format(pdFecsis, gsFormatoFecha)
End If
If Mid(psOpecod, 3, 1) = "2" Then
   oMov.GeneraMovME lnMovNro, psMovNro, pnTCBanco
End If
oMov.ActualizaSaldoMovimiento psMovNro, "+"

oMov.CommitTrans
lbTrans = False
Set oMov = Nothing
GrabaMovimientosCtasIF = 0
Exit Function
GrabaMovimientosCtasIFErr:
        If lbTrans Then
            oMov.RollbackTrans
            lbTrans = False
        End If
        RaiseError Err.Number, Err.Description
End Function
Public Function GrabaMovEfectivo(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
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

oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
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
    oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, Format(pdFechaDoc, gsFormatoFecha)
End If
If pnMovRef <> -1 Then
    'oDMov.ActualizaMov pnMovRef, , gMovEstContabMovContable
    'oDMov.ExtornaMovCtaIF pnMovRef
    oDMov.InsertaMovRef lnMovNro, pnMovRef
End If
If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"
oDMov.CommitTrans
lbTrans = False

GrabaMovEfectivo = 0
Set oDMov = Nothing

Exit Function
GrabaMovEfectivoErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaMovEfectivo", Err.Description & " Error Registro de Movimiento de Efectivo de Caja"
End Function
Public Function GrabaMovCheques(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
                                ByVal pnImporte As Currency, ByVal psAgencia As String, ByVal psIFTpo As String, ByVal psPersCod As String, _
                                ByVal psLsctaIFTpo As String, _
                                ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pdFechaDoc As Date, _
                                ByVal psNroVoucher As String, ByVal pnMoneda As Integer) As Integer
 Dim oDMov As DMov
Set oDMov = New DMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim sSql As String
On Error GoTo GrabaMovGeneralErr

GrabaMovCheques = 1
oDMov.BeginTrans
lbTrans = True

oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
If psNroDoc <> "" Then
    oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, Format(pdFechaDoc, gsFormatoFecha)
End If
If psNroVoucher <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psNroVoucher, Format(pdFechaDoc, gsFormatoFecha)
End If



oDMov.InsertaOpe lnMovNro, pnImporte, psNroDoc, psMovDesc, pnMoneda, psOpecod
oDMov.InsertaMovObj lnMovNro, 1, 1, "11"

oDMov.InsertaMovObjAgenciaArea lnMovNro, 1, 1, psAgencia, ""
oDMov.InsertaMovObjIF lnMovNro, 1, 1, psPersCod, psIFTpo, psLsctaIFTpo


'If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
'    oDMov.GeneraMovME lnMovNro, psMovnro
'End If

oDMov.CommitTrans
lbTrans = False

GrabaMovCheques = 0
Set oDMov = Nothing

Exit Function
GrabaMovGeneralErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaMovEfectivo", Err.Description & " Error Registro de Movimiento de Efectivo de Caja"
                                
                                
End Function


Public Function GrabaMovGeneral(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
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

oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
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
        'oDMov.InsertaMovCtaIF lnMovNro, Mid(psObjetoDeb, 4, 13), Mid(psObjetoDeb, 1, 2), Mid(psObjetoDeb, 18, 10), CGCtaIFConCapital, pnImporte
'        oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
'                                Mid(psObjetoDeb, 4, 13), Mid(psObjetoDeb, 1, 2), Mid(psObjetoDeb, 18, 10)

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
        'oDMov.InsertaMovCtaIF lnMovNro, Mid(psObjetoHab, 4, 13), Mid(psObjetoHab, 1, 2), Mid(psObjetoHab, 18, 10), CGCtaIFConCapital, pnImporte * -1
        
'        oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
'                                Mid(psObjetoHab, 4, 13), Mid(psObjetoHab, 1, 2), Mid(psObjetoHab, 18, 10)

    Case ObjDescomEfectivo
        oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, psObjetoHab
    Case ObjPersona
        oDMov.InsertaMovGasto lnMovNro, psObjetoHab, ""
    Case Else
End Select
End If
If psNroDoc <> "" Then
    oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, Format(pdFechaDoc, gsFormatoFecha)
End If
If psNroVoucher <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psNroVoucher, Format(pdFechaDoc, gsFormatoFecha)
End If
If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If

oDMov.ActualizaSaldoMovimiento psMovNro, "+"
oDMov.CommitTrans
lbTrans = False

GrabaMovGeneral = 0
Set oDMov = Nothing

Exit Function
GrabaMovGeneralErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaMovEfectivo", Err.Description & " Error Registro de Movimiento de Efectivo de Caja"
End Function
Public Function GetRetirosEfectivoCaja(ByVal psOpecod As String, ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset
Dim lsFiltroAge As String
sql = " SELECT  DOCMOV.cDocAbrev, DOCMOV.cDocNro, VOUC.cDocNro AS VOUCHER, P.cPersNombre, " _
    & "         RTRIM(C1.cCtaIFDesc) + ' ' + C.cCtaIFDesc AS CUENTA, " _
    & "         ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE))  AS IMPORTE, M.NMOVNRO, M.CMOVDESC, MIF.cPersCod, MIF.cIfTpo, MIF.cCtaIFCod " _
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
    & "         JOIN CTAIF C1 ON SUBSTRING(MIF.cCtaIFCod,1,LEN(C1.cCtaIFCod))=C1.cCtaIFCod AND LEN(C1.cCtaIFCod)<=5 " _
    & "                             AND MIF.CPERSCOD=C1.CPERSCOD " _
    & "         JOIN CTAIF C ON MIF.cCtaIFCod=C.cCtaIFCod AND MIF.CPERSCOD=C.CPERSCOD " _
    & " WHERE   M.COPECOD ='" & psOpecod & "' AND M.nMovFlag = '" & gMovFlagVigente & "' " _
    & "         AND SUBSTRING(M.cMovnro,1,8 ) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "'  " _
    & "         AND NOT EXISTS (SELECT  M1.CMOVNRO FROM  MOV M1 " _
    & "                                 JOIN MOVREF MR ON M1.nMOVNRO =MR.nMOVNRO " _
    & "                         WHERE MR.nMOVNROREF = M.nMOVNRO AND M1.nMovFlag = " & gMovFlagVigente & ") " _
    & " ORDER BY M.CMOVNRO "

Set rs = oConect.CargaRecordSet(sql)
Set GetRetirosEfectivoCaja = rs

End Function
Public Function GrabaDepositoCheques(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
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

oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
lnMovItem = lnMovItem + 1

oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, pnTotalImporte
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psObjetoIF, 4, 13), Mid(psObjetoIF, 1, 2), Mid(psObjetoIF, 18, 10)

'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
'oDMov.InsertaMovCtaIF lnMovNro, Mid(psObjetoIF, 4, 13), Mid(psObjetoIF, 1, 2), Mid(psObjetoIF, 18, 10), CGCtaIFConCapital, pnTotalImporte
'oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                            Mid(psObjetoIF, 4, 13), Mid(psObjetoIF, 1, 2), Mid(psObjetoIF, 18, 10)


lnMovOrden = 0
If Not rs Is Nothing Then
    If rs.State = adStateOpen Then
        If Not rs.EOF And Not rs.BOF Then
            Do While Not rs.EOF
                If Val(rs!Ok) = 1 Then
                    lsFiltro = ""
                    lsFiltro = oCont.GetFiltroObjetos(ObjProductosCMACT, psCtaHaber, rs!cObjetoCod, False)
                    
                    lsFiltro = lsFiltro & oCont.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaHaber, rs!AreaCod & rs!AgeCod, False)
                    
                    
                    lnMovItem = lnMovItem + 1
                    lnMovOrden = 0
                    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber + lsFiltro, rs!Importe * -1
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rs!cObjetoCod
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
                    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, rs!AgeCod, rs!AreaCod
                    
                    oDMov.ActualizaCheque TpoDocCheque, rs![N° Cheque], rs!cPerscod, rs!cIFTpo, , , , rs![Cta Banco], , gCGEstadosChqDepositado
                    
                    oDMov.InsertaMovRef lnMovNro, rs!nMovNro
                End If
                rs.MoveNext
            Loop
        End If
    End If
End If
'If psNroDoc <> "" Then
'    oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, Format(pdFechaDoc, gsFormatoFecha)
'End If
If Mid(psOpecod, 3, 1) = "2" Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"

oDMov.CommitTrans
lbTrans = False

GrabaDepositoCheques = 0
Set oDMov = Nothing

Exit Function
GrabaDepositoChequesErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaDepositoCheques", Err.Description & " Error Registro de Movimiento de Efectivo de Caja"
End Function

Public Function GrabaDevolucionCheques(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
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
On Error GoTo GrabaDevolucionChequesErr

GrabaDevolucionCheques = 1
Set oCont = New NContFunciones
oDMov.BeginTrans
lbTrans = True
lnMovItem = 0
lnMovOrden = 0
oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
lnMovItem = lnMovItem + 1

'Grabamos la cuenta de devolucion
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, pnTotalImporte

lnMovOrden = 0
If Not rs Is Nothing Then
    If rs.State = adStateOpen Then
        If Not rs.EOF And Not rs.BOF Then
            Do While Not rs.EOF
                If Val(rs!Ok) = 1 Then
                    lsFiltro = ""
                    lsFiltro = oCont.GetFiltroObjetos(ObjProductosCMACT, psCtaHaber, rs!cObjetoCod, False)
                    lsFiltro = lsFiltro & oCont.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaHaber, rs!AreaCod & rs!AgeCod, False)
                    lnMovItem = lnMovItem + 1
                    lnMovOrden = 0
                    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber + lsFiltro, rs!Importe * -1
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rs!cObjetoCod
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
                    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, rs!AgeCod, rs!AreaCod
                    
                    oDMov.ActualizaCheque TpoDocCheque, rs![N° Cheque], rs!cPerscod, rs!cIFTpo, , , , rs![Cta Banco], , gCGEstadosChqRechazado, , , gChqEstRechazado
                    
                    oDMov.ActualizaMov rs!nMovNro, , gMovFlagEliminado
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
If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"
oDMov.CommitTrans
lbTrans = False

GrabaDevolucionCheques = 0
Set oDMov = Nothing

Exit Function
GrabaDevolucionChequesErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaDevolucionCheques", Err.Description & " Error Registro de Movimiento de Efectivo de Caja"
End Function

Public Function GetMotivosDepRet(ByVal pnTipoMot As MotivosDepRet) As ADODB.Recordset
sql = " SELECT cMotDesc, cMotCod  " _
    & " FROM MOTIVOSDEPRET WHERE cMotCod LIKE '" & pnTipoMot & "%' AND LEN(CMOTCOD)=3"

Set rs = oConect.CargaRecordSet(sql)
Set GetMotivosDepRet = rs
End Function

Public Function GrabaDepRetBancosVarios(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
                                        ByVal psCtaContBanco As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                        ByVal pnDocTpo As TpoDoc, ByVal psDocNro As String, ByVal psFechaDoc As String, _
                                        ByVal psDocNroVoucher As String, ByVal rsObjHaber As ADODB.Recordset, _
                                        ByVal lnTpoObjIf As TpoObjetos, ByVal psCtaIf As String, ByVal pbDeposito As Boolean, _
                                        Optional pnMotivoNANC As Long = -1, Optional psMotObjPadre As String = "", Optional psMotObjeto As String = "", Optional ByVal psctaAho As String = "", _
                                        Optional ByVal pnDocTpoNANCAux As TpoDoc = -1, Optional ByVal psNroDocNANCAux As String = "", Optional ByVal pnImporteAux As Currency, _
                                        Optional pnMotivoNANCAux As Long = -1, Optional psMotObjPadreAux As String = "", Optional psMotObjetoAux As String = "", Optional ByVal psCtaAhoAux As String = "", _
                                        Optional ByVal pbGrabaNegocio As Boolean = False, Optional ByVal pbBitCentral As Boolean = False, _
                                        Optional psOpeCodNegocio As String = "", _
                                        Optional prsNego As ADODB.Recordset, Optional psAjusteCta As String = "", Optional pnAjusteMonto As Currency = 0, Optional pnTCBanco As Currency = 0) As Integer
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
Dim lsMovNroNegocio As String
Set oCont = New NContFunciones
On Error GoTo ErrorGrabaDepRetBancosVarios

GrabaDepRetBancosVarios = 1
oDMov.BeginTrans
lbTrans = True

oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1
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

oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber + lsSubCta, (pnImporte - pnAjusteMonto) * IIf(pbDeposito, -1, 1)

If psCtaHaber = "111109" Or psCtaHaber = "112109" Then
    
    oDMov.InsertaMovHabilitaCaja lnMovNro, Abs((pnImporte - pnAjusteMonto)), "025", "026", lsSubCta, Mid(psCtaHaber, 3, 1)
    
    If psCtaHaber = "111109" Then
        oDMov.InsertaMovObjEfectivo lnMovNro, 1, 1, "1M005"
    Else
        oDMov.InsertaMovObjEfectivo lnMovNro, 1, 1, "2B005"
    End If
End If

If Not rsObjHaber Is Nothing Then
    If rsObjHaber.State = adStateOpen Then
        rsObjHaber.MoveFirst
        Do While Not rsObjHaber.EOF
            lnMovOrden = lnMovOrden + 1
            Select Case rsObjHaber!cObjetoCod
                Case ObjCMACAgenciaArea
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!cObjetoCod
                    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(rsObjHaber!objeto, 4, 2), Mid(rsObjHaber!objeto, 1, 3)
                Case ObjCMACAgencias
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!cObjetoCod
                    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!objeto, "013"
                Case ObjDescomEfectivo
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!cObjetoCod
                    oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!objeto
                Case ObjEntidadesFinancieras
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!cObjetoCod
                    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(rsObjHaber!objeto, 4, 13), Mid(rsObjHaber!objeto, 1, 2), Mid(rsObjHaber!objeto, 18, 10)
                Case Else
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!objeto
            End Select
            rsObjHaber.MoveNext
        Loop
    End If
End If

If pnAjusteMonto <> 0 Then
   oDMov.InsertaMovCta lnMovNro, lnMovItem + 1, psAjusteCta, pnAjusteMonto * IIf(pbDeposito, -1, 1)
End If

Select Case Val(pnDocTpo)
    Case TpoDocNotaAbono, TpoDocNotaCargo
        oDMov.InsertaNotaAbonoCargo pnDocTpo, psDocNro, gNCNAEnMovimiento, 0, pnImporte, psCtaIf, ObjEntidadesFinancieras, Mid(psOpecod, 3, 1)
'        oDMov.ActualizaNotaAbonoCargo pnDocTpo, psDocNro, gNCNAEnMovimiento
End Select

Select Case Val(pnDocTpo)
    Case TpoDocCarta, TpoDocCheque, TpoDocNotaCargo
        If pbGrabaNegocio Then
            Dim oCapta As NCapMovimientos
            Dim lnSaldo As Double
            Set oCapta = New NCapMovimientos
            Dim oCon As New DConecta
            Dim oDis As New NRHProcesosCierre
            
            If pbBitCentral Then
               lsMovNroNegocio = oDMov.GeneraMovNro(, , , psMovNro)
               lnSaldo = oCapta.CapCargoCuentaAho(psCtaAhoAux, pnImporteAux, psOpeCodNegocio, lsMovNroNegocio, psMovDesc, TpoDocNotaCargo, psDocNro, Mid(psCtaIf, 4, 13), , , , , , , , , True, "", , oDMov.GetConexion, False)
            Else
                If oCon.AbreConexion Then 'Remota(Left(psCtaAhoAux, 2))
                    oCon.BeginTrans
                    lnSaldo = oDis.Cargo(psCtaAhoAux, pnImporteAux, psOpeCodNegocio, "112" & Left(psCtaAhoAux, 2), Right(psMovNro, 4), psDocNro, "CARGO : " & psMovDesc, oCon, GetFechaMov(psMovNro, True))
                    oCon.CommitTrans
                End If
                oCon.CierraConexion
                Set oCon = Nothing
                Set oDis = Nothing
            End If
        End If
End Select

If psDocNro <> "" Then
    oDMov.InsertaMovDoc lnMovNro, pnDocTpo, psDocNro, Format(CDate(psFechaDoc), gsFormatoFecha)
End If
If psDocNroVoucher <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocNroVoucher, Format(CDate(psFechaDoc), gsFormatoFecha)
End If
If pnDocTpoNANCAux <> -1 And psNroDocNANCAux <> "" Then
    oDMov.InsertaNotaAbonoCargo pnDocTpoNANCAux, psNroDocNANCAux, gNCNAEnMovimiento, 0, pnImporteAux, psCtaIf, ObjEntidadesFinancieras, Mid(psOpecod, 3, 1)
    oDMov.InsertaMovDoc lnMovNro, pnDocTpoNANCAux, psNroDocNANCAux, Format(CDate(psFechaDoc), gsFormatoFecha)
End If

If Not prsNego Is Nothing Then
    If Not prsNego.State = adStateClosed Then
        Dim oNeg As New NNegOpePendientes
        Do While Not prsNego.EOF
            If prsNego.Fields(1).value = "1" Then   'Check Seleccionado
                'Actualizamos referencia en Negocio
                If Not pbBitCentral Then
                    oNeg.ActualizaNegocioReferencia prsNego.Fields(10), prsNego.Fields(7), psMovNro, prsNego.Fields(11)
                End If
                oDMov.InsertaMovRef lnMovNro, prsNego.Fields(10), Right(prsNego.Fields(7), 2)
            End If
            prsNego.MoveNext
        Loop
        Set oNeg = Nothing
    End If
Else
    Dim lsCtaCaracterOut As String
    If oDMov.CuentaEsPendiente(psCtaHaber + lsSubCta, lsCtaCaracterOut, IIf(pbDeposito, "A", "D")) Then
        If (lsCtaCaracterOut = "A" And pbDeposito) Or (lsCtaCaracterOut = "D" And Not pbDeposito) Then
            oDMov.InsertaMovPendientesRend lnMovNro, psCtaHaber + lsSubCta, pnImporte
        End If
    End If
End If

GrabaDepRetBancosVarios = 0
If Mid(psOpecod, 3, 1) = 2 Then
    oDMov.GeneraMovME lnMovNro, psMovNro, pnTCBanco
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"


oDMov.CommitTrans
lbTrans = False
Set oDMov = Nothing
Exit Function
ErrorGrabaDepRetBancosVarios:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaRendicionGiroDocumento", lsMsgErr
End Function

Public Function GrabaAperturaCtaIF(ByVal psMovNro As String, ByVal psOpecod As String, ByVal pdFecha As Date, ByVal psMovDesc As String, pnImporte As Currency, _
       ByVal prsCtaApe As ADODB.Recordset, pbCreaSubCta As Boolean, psSubCtaIF As String, psSubCtaIFDesc As String, _
       ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, ByVal rsBillete As ADODB.Recordset, ByVal rsMoneda As ADODB.Recordset, _
       ByVal psNroCartaApert As String, ByVal pdFecsis As Date, _
       ByVal psEntTransferIFCod As String, pnEntTransferIFImporte As Currency, _
       ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pdFechaDoc As Date, ByVal psNroVoucher As String, _
       ByVal prsCheque As ADODB.Recordset, ByVal prsOtro As ADODB.Recordset, ByVal prsObj As ADODB.Recordset, _
       Optional psTpoPF As Integer, Optional ByVal lTasaInteres As Variant) As Integer
       'RIRO20140530 ERS017, Se agregó lTasaInteres
            
Dim oContF     As NContFunciones
Dim oDMov      As DMov
Dim oOpe       As DOperacion
Dim oCont      As NContFunciones

Dim lnMovItem  As Long
Dim lnMovOrden As Long
Dim lnMovNro   As Long
Dim lsMovNro   As String
Dim lsSubCta   As String
Dim lbTrans    As Boolean
Dim lsCuenta   As String
Dim lsFiltro   As String

On Error GoTo GrabaAperturaCtaIFErr

Set oDMov = New DMov
Set oContF = New NContFunciones
Set oOpe = New DOperacion
Set oCont = New NContFunciones

GrabaAperturaCtaIF = 1

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"

'Guardamos la cuenta en el debe
lsCuenta = oOpe.EmiteOpeCta(psOpecod, "D", "0", Format(psIFTpo, "00") & psPersCod & psSubCtaIF, ObjEntidadesFinancieras)
If pbCreaSubCta Then
    oDMov.InsertaCuentaIF psPersCod, psIFTpo, psSubCtaIF, psSubCtaIFDesc, pdFecha, Format(pdFecha, gsFormatoFecha), Format(pdFecha, gsFormatoFecha), 0, pdFecha, gEstadoCtaIFActiva, psMovNro, , psTpoPF
End If
lnMovItem = 0

Do While Not prsCtaApe.EOF

    If prsCtaApe!Codigo <> "" And prsCtaApe!Importe <> 0 Then
    'verificando que si ingreso mnto en euros, entonces el check sea = 1
    'edpyme
        If prsCtaApe!MontoEuros > 0 And prsCtaApe!Check = 0 Then
            prsCtaApe!Check = 1
        End If
    
        ' RIRO20140530 ERS017, Se agregó lTasaInteres(lnMovItem)
        lnMovItem = lnMovItem + 1: lnMovOrden = 1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta, prsCtaApe!Importe
        oDMov.InsertaCuentaIF psPersCod, psIFTpo, prsCtaApe!Codigo, prsCtaApe!Cuenta_Nro, pdFecha, "", "", nVal(prsCtaApe!Plazo), pdFecha + nVal(prsCtaApe!Plazo), _
                             gEstadoCtaIFRegistrada, psMovNro, IIf(prsCtaApe!Check = 1, prsCtaApe!MontoEuros, 0), psTpoPF
        oDMov.InsertaCuentaIFInteres psPersCod, psIFTpo, prsCtaApe!Codigo, pdFecha, nVal(prsCtaApe!PerTasa), lTasaInteres(lnMovItem), psMovNro 'nVal(prsCtaApe![T.Int.Efect]), psMovNro
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, prsCtaApe!Codigo
    End If
    prsCtaApe.MoveNext
Loop

lsCuenta = oOpe.EmiteOpeCta(psOpecod, "H", "0")
If lsCuenta = "" Then
    Err.Raise 50001, "nCajaGeneral: GrabaAperturaCtaIF", "Cuenta Contable de Efectivo vacia. Consultar con Sistemas"
End If
oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsBillete, lsCuenta, "H"
oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsMoneda, lsCuenta, "H"

'Transferencia
If psEntTransferIFCod <> "" And pnEntTransferIFImporte > 0 Then
    lnMovItem = lnMovItem + 1: lnMovOrden = 0
    
    '*** PEAC 20110825
    'lsCuenta = oOpe.EmiteOpeCta(psOpeCod, "H", "1", psEntTransferIFCod, ObjEntidadesFinancieras) '' si es el BCR DEBE MANDAR A LA CUENTA 111201
    lsCuenta = oOpe.EmiteOpeCta(psOpecod, "H", IIf(psOpecod = "401514" And Mid(psEntTransferIFCod, 4, 13) = "1090100822183", "3", "1"), psEntTransferIFCod, ObjEntidadesFinancieras)
    
    If lsCuenta = "" Then
        Err.Raise 50001, "nCajaGeneral: GrabaAperturaCtaIF", "Cuenta Contable de Inst.Financ Origen vacia. Consultar con Sistemas"
    End If
    oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta, pnEntTransferIFImporte * -1
    lnMovOrden = lnMovOrden + 1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psEntTransferIFCod, 4, 13), Mid(psEntTransferIFCod, 1, 2), Mid(psEntTransferIFCod, 18, 10)
End If

'Cheques recibidos - Cheque de Gerencia
If Not prsCheque Is Nothing Then
    If prsCheque.State = adStateOpen Then
        lsCuenta = oOpe.EmiteOpeCta(psOpecod, "H", "2")
        If lsCuenta = "" Then
            Err.Raise 50001, "nCajaGeneral: GrabaAperturaCtaIF", "Cuenta Contable de Cheques Recibidos vacio. Consultar con Sistemas"
        End If
        Do While Not prsCheque.EOF
            If prsCheque!Opc = "1" Then
                lnMovItem = lnMovItem + 1: lnMovOrden = 0
                
                lsFiltro = oCont.GetFiltroObjetos(ObjProductosCMACT, lsCuenta, prsCheque!Cuenta, False)
                lsFiltro = lsFiltro & oCont.GetFiltroObjetos(ObjCMACAgenciaArea, lsCuenta, prsCheque!cAreaCod & prsCheque!cAgeCod, False)
                oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta & lsFiltro, prsCheque!Importe * -1
                
                lnMovOrden = lnMovOrden + 1
                oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, prsCheque!Cuenta
                lnMovOrden = lnMovOrden + 1
                oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
                oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, prsCheque!cAgeCod, prsCheque!cAreaCod
                
                oDMov.ActualizaCheque TpoDocCheque, prsCheque!NroCheque, prsCheque!cPerscod, prsCheque!cIFTpo, , , , , , gCGEstadosChqDepositado
                oDMov.InsertaMovRef lnMovNro, prsCheque!nMovNro
            End If
            prsCheque.MoveNext
        Loop
    End If
End If

'Otras Formas de Pago
oDMov.GrabaMovOtrosCtaObj lnMovNro, lnMovItem, prsOtro, prsObj, "H"

If psNroDoc <> "" Then
    oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, Format(pdFechaDoc, gsFormatoFecha)
End If
If psNroVoucher <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psNroVoucher, Format(pdFechaDoc, gsFormatoFecha)
End If
If psNroCartaApert <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocCarta, psNroCartaApert, Format(pdFecha, gsFormatoFecha)
End If

If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If

oDMov.ActualizaSaldoMovimiento psMovNro, "+"

oDMov.CommitTrans
lbTrans = False
GrabaAperturaCtaIF = 0
Set oDMov = Nothing
Set oContF = Nothing

Exit Function
GrabaAperturaCtaIFErr:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaAperturaCtaIF", lsMsgErr & " Error Graba Apertura de Ctas IF"
End Function

Public Function GrabaAdeudadoProyectado(ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, _
          ByVal psCuentaCod As String, ByVal psCuentaDesc As String, _
          ByVal pdFecha As Date, pbCreaSubCta As Boolean, psSubCtaIF As String, psSubCtaIFDesc As String, _
          ByVal pnPeriodo As Integer, pnInteres As Currency, _
          pnCapital As Currency, pnNroCuotas As Integer, _
          pnGracia As Currency, pnComisionIni As Currency, pnComisionMonto As Currency, _
          psInterno As String, pnCuotaPagok As Integer, _
          pnMonedaPago As Integer, pnTpoCuota As CGAdeudCalTpoCuota, pnTramo As Currency, _
          ByVal pnComisionCuota As Currency, psLinCredCod As String, _
          pnFechaFija As Integer, ByVal pnPlazo As Currency, prsAdeud As ADODB.Recordset, Optional psMovNro As String = "") As Integer
Dim oDMov      As DMov
Dim lbTrans    As Boolean
On Error GoTo GrabaAdeudadoProyectadoErr
Set oDMov = New DMov
oDMov.BeginTrans
lbTrans = True
If pbCreaSubCta Then
    oDMov.InsertaCuentaIF psPersCod, psIFTpo, psSubCtaIF, psSubCtaIFDesc, pdFecha, "", "", 0, pdFecha, gEstadoCtaIFActiva, psMovNro
End If

oDMov.InsertaCuentaIF psPersCod, psIFTpo, psCuentaCod, psCuentaDesc, pdFecha, Format(pdFecha, gsFormatoFecha), Format(pdFecha, gsFormatoFecha), pnPlazo, pdFecha, _
                     gEstadoCtaIFRegistrada, psMovNro
oDMov.InsertaCuentaIFInteres psPersCod, psIFTpo, psCuentaCod, pdFecha, pnPeriodo, pnInteres, psMovNro

If Not prsAdeud Is Nothing Then
    If Not prsAdeud.State = adStateClosed Then
        oDMov.InsertaDatosAdeudados prsAdeud, psPersCod, psIFTpo, psCuentaCod, pnCapital, pnNroCuotas, pnGracia, pnComisionIni, pnComisionMonto, psInterno, pnCuotaPagok, Format(pdFecha, "YYYY/MM/DD"), pnMonedaPago, Format(pdFecha, "YYYY/MM/DD"), pnTpoCuota, pnTramo, pnComisionCuota, psLinCredCod, pnFechaFija, psMovNro
    Else
        Err.Raise 50001, "GrabaAdeudadoProyectado", "Falta Calendario de Adeudado"
    End If
Else
    Err.Raise 50001, "GrabaAdeudadoProyectado", "Falta Calendario de Adeudado"
End If
oDMov.CommitTrans
lbTrans = False
Exit Function
GrabaAdeudadoProyectadoErr:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
    End If
    Err.Raise Err.Number, Err.Source, lsMsgErr
End Function

'' Nuevo Pepe Trujillo
Public Function GrabaRegPagareAdeudado(ByVal psMovNro As String, ByVal psOpecod As String, ByVal pdFecha As Date, ByVal psMovDesc As String, ByVal pnImporte As Currency, ByVal pnCapitalVac As Currency, _
       ByVal psCuentaCod As String, psCuentaDesc As String, pbCreaSubCta As Boolean, psSubCtaIF As String, psSubCtaIFDesc As String, _
       ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, ByVal rsBillete As ADODB.Recordset, ByVal rsMoneda As ADODB.Recordset, _
       ByVal psNroCartaApert As String, ByVal pdFecsis As Date, _
       ByVal psEntTransferIFCod As String, pnEntTransferIFImporte As Currency, _
       ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pdFechaDoc As Date, ByVal psNroVoucher As String, _
       ByVal prsCheque As ADODB.Recordset, ByVal prsOtro As ADODB.Recordset, ByVal prsObj As ADODB.Recordset, _
       ByVal prsAdeud As ADODB.Recordset, pnGracia As Currency, pnComisionIni As Currency, pnComisionMonto As Currency, psInterno As String, pnCuotaPagok As Integer, pnNroCuotas As Integer, _
       ByVal pnPlazo As Currency, pnPeriodo As Currency, pnInteres As Currency, pnTpoCuota As CGAdeudCalTpoCuota, pnTramo As Currency, pnMonedaPago As Integer, _
       ByVal pnComisionCuota As Currency, psLinCredCod As String, pnFechaFija As Integer, ByVal pnPorcentAfect As Currency, _
       Optional psCtaOrdenD As String = "", Optional psCtaOrdenH As String = "", _
       Optional pnMontoPresConce As Currency = 0, Optional pbDatosContables As Integer = 1) As Integer

Dim oContF     As NContFunciones
Dim oDMov      As DMov
Dim oOpe       As DOperacion
Dim oCont      As NContFunciones

Dim lnMovItem  As Long
Dim lnMovOrden As Long
Dim lnMovNro   As Long
Dim lsMovNro   As String
Dim lsSubCta   As String
Dim lbTrans    As Boolean
Dim lsCuenta   As String
Dim lsFiltro   As String

On Error GoTo GrabaRegPagareAdeudadoErr

Set oDMov = New DMov
Set oContF = New NContFunciones
Set oOpe = New DOperacion
Set oCont = New NContFunciones

GrabaRegPagareAdeudado = 1

oDMov.BeginTrans
lbTrans = True
If pbDatosContables = 1 Then
    oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable
    lnMovNro = oDMov.GetnMovNro(psMovNro)
    oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
Else
    oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, 13
    lnMovNro = oDMov.GetnMovNro(psMovNro)
    oDMov.InsertaMovContMIVIVIENDA lnMovNro, pnImporte, "0", "0"
End If
'Guardamos la cuenta en el Haber
lsCuenta = oOpe.EmiteOpeCta(psOpecod, "H", "0")
If pbCreaSubCta Then
    oDMov.InsertaCuentaIF psPersCod, psIFTpo, psSubCtaIF, psSubCtaIFDesc, pdFecha, "", "", 0, pdFecha, gEstadoCtaIFActiva, psMovNro, , , pbDatosContables, lnMovNro
End If
lnMovItem = 0

'ALPA 20150206*******************************************
If pbDatosContables = 1 Then
    lnMovItem = lnMovItem + 1: lnMovOrden = 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta, pnImporte * -1
    oDMov.InsertaCuentaIF psPersCod, psIFTpo, psCuentaCod, psCuentaDesc, pdFecha, Format(pdFecha, gsFormatoFecha), Format(pdFecha, gsFormatoFecha), pnPlazo, pdFecha, _
                         gEstadoCtaIFRegistrada, psMovNro, , , pbDatosContables, lnMovNro
    
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCuentaCod
    
    If pnComisionMonto <> 0 Then   'Comision Inicial
        lnMovItem = lnMovItem + 1

        oDMov.InsertaMovCta lnMovNro, lnMovItem, "19" & Mid(psOpecod, 3, 1) & "10501", pnComisionMonto 'EJVG20130306 Cambio x NIIFs
    End If
Else
    lnMovItem = lnMovItem + 1: lnMovOrden = 1
    oDMov.InsertaMovCtaMIVIVIENDA lnMovNro, lnMovItem, lsCuenta, pnImporte * -1
    oDMov.InsertaCuentaIF psPersCod, psIFTpo, psCuentaCod, psCuentaDesc, pdFecha, Format(pdFecha, gsFormatoFecha), Format(pdFecha, gsFormatoFecha), pnPlazo, pdFecha, _
                         gEstadoCtaIFRegistrada, psMovNro, , , pbDatosContables, lnMovNro
    
    oDMov.InsertaMovObjMIVIVIENDA lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIFMIVIVIENDA lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCuentaCod
    
    If pnComisionMonto <> 0 Then   'Comision Inicial
        lnMovItem = lnMovItem + 1

        oDMov.InsertaMovCtaMIVIVIENDA lnMovNro, lnMovItem, "19" & Mid(psOpecod, 3, 1) & "10501", pnComisionMonto 'EJVG20130306 Cambio x NIIFs
    End If
End If
'*********************************************************
If Not prsAdeud Is Nothing Then
    If Not prsAdeud.State = adStateClosed Then
        'ALPA20130719***************
        'oDMov.InsertaDatosAdeudados prsAdeud, psPersCod, psIFTpo, psCuentaCod, pnImporte, pnNroCuotas, pnGracia, pnComisionIni, pnComisionMonto, psInterno, pnCuotaPagok, Format(pdFecha, "YYYY/MM/DD"), pnMonedaPago, Format(pdFecha, "YYYY/MM/DD"), pnTpoCuota, pnTramo, pnComisionCuota, psLinCredCod, pnFechaFija, lsMovNro, , pnCapitalVac, , pnPorcentAfect
        oDMov.InsertaDatosAdeudados prsAdeud, psPersCod, psIFTpo, psCuentaCod, pnImporte, pnNroCuotas, pnGracia, pnComisionIni, pnComisionMonto, psInterno, pnCuotaPagok, Format(pdFecha, "YYYY/MM/DD"), pnMonedaPago, Format(pdFecha, "YYYY/MM/DD"), pnTpoCuota, pnTramo, pnComisionCuota, psLinCredCod, pnFechaFija, lsMovNro, , pnCapitalVac, , pnPorcentAfect, pnMontoPresConce
        '***************************
    End If
End If

'Interes
oDMov.CGActualizaCuentaIFInteres psPersCod, Format(psIFTpo, "00"), psCuentaCod, pdFecha, pnPeriodo, pnInteres, psMovNro

'Actualiza Saldos de Linea de Credito
oDMov.ActualizaColocAdeudadoSaldo psLinCredCod, GetFechaMov(psMovNro, True), pnImporte


lsCuenta = oOpe.EmiteOpeCta(psOpecod, "D", "0")
'ALPA 20150206*******************************
If pbDatosContables = 1 Then
    oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsBillete, lsCuenta, "D"
    oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsMoneda, lsCuenta, "D"
Else
    oDMov.GrabaMovimientoEfectivoMIVIVIENDA lnMovNro, lnMovItem, rsBillete, lsCuenta, "D"
    oDMov.GrabaMovimientoEfectivoMIVIVIENDA lnMovNro, lnMovItem, rsMoneda, lsCuenta, "D"
End If
'********************************************
'Transferencia
'ALPA 20150206*******************************
If pbDatosContables = 1 Then
    If psEntTransferIFCod <> "" And pnEntTransferIFImporte > 0 Then
        lnMovItem = lnMovItem + 1: lnMovOrden = 0
        lsCuenta = oOpe.EmiteOpeCta(psOpecod, "D", "1", psEntTransferIFCod, Left(psEntTransferIFCod, 2), True, False)
        oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta, pnEntTransferIFImporte
        lnMovOrden = lnMovOrden + 1
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psEntTransferIFCod, 4, 13), Mid(psEntTransferIFCod, 1, 2), Mid(psEntTransferIFCod, 18, 10)
    End If
Else
    If psEntTransferIFCod <> "" And pnEntTransferIFImporte > 0 Then
        lnMovItem = lnMovItem + 1: lnMovOrden = 0
        lsCuenta = oOpe.EmiteOpeCta(psOpecod, "D", "1", psEntTransferIFCod, Left(psEntTransferIFCod, 2), True, False)
        oDMov.InsertaMovCtaMIVIVIENDA lnMovNro, lnMovItem, lsCuenta, pnEntTransferIFImporte
        lnMovOrden = lnMovOrden + 1
        oDMov.InsertaMovObjMIVIVIENDA lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIFMIVIVIENDA lnMovNro, lnMovItem, lnMovOrden, Mid(psEntTransferIFCod, 4, 13), Mid(psEntTransferIFCod, 1, 2), Mid(psEntTransferIFCod, 18, 10)
    End If
End If
'End If
'Cheques recibidos - Cheque de Gerencia

If Not prsCheque Is Nothing Then
    If prsCheque.State = adStateOpen Then
        lsCuenta = oOpe.EmiteOpeCta(psOpecod, "D", "2")
        Do While Not prsCheque.EOF
            If prsCheque!Opc = "1" Then
                If pbDatosContables = 1 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
    
                    lsFiltro = oCont.GetFiltroObjetos(ObjProductosCMACT, lsCuenta, prsCheque!Cuenta, False)
                    lsFiltro = lsFiltro & oCont.GetFiltroObjetos(ObjCMACAgenciaArea, lsCuenta, prsCheque!cAreaCod & prsCheque!cAgeCod, False)
                    oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta & lsFiltro, prsCheque!Importe
    
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, prsCheque!Cuenta
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
                    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, prsCheque!cAgeCod, prsCheque!cAreaCod
    
                    oDMov.ActualizaCheque TpoDocCheque, prsCheque!NroCheque, prsCheque!cPerscod, prsCheque!cIFTpo, , , , , , gCGEstadosChqDepositado
                    oDMov.InsertaMovRef lnMovNro, prsCheque!nMovNro
                Else
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
    
                    lsFiltro = oCont.GetFiltroObjetos(ObjProductosCMACT, lsCuenta, prsCheque!Cuenta, False)
                    lsFiltro = lsFiltro & oCont.GetFiltroObjetos(ObjCMACAgenciaArea, lsCuenta, prsCheque!cAreaCod & prsCheque!cAgeCod, False)
                    oDMov.InsertaMovCtaMIVIVIENDA lnMovNro, lnMovItem, lsCuenta & lsFiltro, prsCheque!Importe
    
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObjMIVIVIENDA lnMovNro, lnMovItem, lnMovOrden, prsCheque!Cuenta
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObjMIVIVIENDA lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
                    oDMov.InsertaMovObjAgenciaAreaMIVIVIENDA lnMovNro, lnMovItem, lnMovOrden, prsCheque!cAgeCod, prsCheque!cAreaCod
    
                    oDMov.ActualizaCheque TpoDocCheque, prsCheque!NroCheque, prsCheque!cPerscod, prsCheque!cIFTpo, , , , , , gCGEstadosChqDepositado
                    oDMov.InsertaMovRefMIVIVIENDA lnMovNro, prsCheque!nMovNro

                End If
            End If
            prsCheque.MoveNext
        Loop
    End If
End If

'Otras Formas de Pago
'ALPA 20150206*********
If pbDatosContables = 1 Then
    oDMov.GrabaMovOtrosCtaObj lnMovNro, lnMovItem, prsOtro, prsObj, "D"
Else
    oDMov.GrabaMovOtrosCtaObjMIVIVIENDA lnMovNro, lnMovItem, prsOtro, prsObj, "D"
End If
'**********************
'ALPA 20150206*********
If pbDatosContables = 1 Then
    If psCtaOrdenD <> "" Then 'Cuentas de Orden - Pepe
        lnMovItem = lnMovItem + 1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaOrdenD, pnImporte
        lnMovItem = lnMovItem + 1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaOrdenH, pnImporte * -1
    End If
    
    If psNroDoc <> "" Then
        oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, Format(pdFechaDoc, gsFormatoFecha)
    End If
    If psNroVoucher <> "" Then
        oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psNroVoucher, Format(pdFechaDoc, gsFormatoFecha)
    End If
    If psNroCartaApert <> "" Then
        oDMov.InsertaMovDoc lnMovNro, TpoDocCarta, psNroCartaApert, Format(pdFecha, gsFormatoFecha)
    End If
    
    If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
        oDMov.GeneraMovME lnMovNro, psMovNro
    End If
Else
    If psCtaOrdenD <> "" Then 'Cuentas de Orden - Pepe
        lnMovItem = lnMovItem + 1
        oDMov.InsertaMovCtaMIVIVIENDA lnMovNro, lnMovItem, psCtaOrdenD, pnImporte
        lnMovItem = lnMovItem + 1
        oDMov.InsertaMovCtaMIVIVIENDA lnMovNro, lnMovItem, psCtaOrdenH, pnImporte * -1
    End If
    
    If psNroDoc <> "" Then
        oDMov.InsertaMovDocMIVIVIENDA lnMovNro, pnTpoDoc, psNroDoc, Format(pdFechaDoc, gsFormatoFecha)
    End If
    If psNroVoucher <> "" Then
        oDMov.InsertaMovDocMIVIVIENDA lnMovNro, TpoDocVoucherEgreso, psNroVoucher, Format(pdFechaDoc, gsFormatoFecha)
    End If
    If psNroCartaApert <> "" Then
        oDMov.InsertaMovDocMIVIVIENDA lnMovNro, TpoDocCarta, psNroCartaApert, Format(pdFecha, gsFormatoFecha)
    End If
    
'    If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
'        oDMov.GeneraMovME lnMovNro, psMovNro
'    End If
End If
'Estadistica para control
If Not (pnMonedaPago = "2" And Mid(Mid(psEntTransferIFCod, 18, 10), 3, 1) = "1") Then
    oDMov.GrabaCtaIFEstadistica psPersCod, CStr(Format(psIFTpo, "00")), psCuentaCod, "0", 0, lnMovNro, 1, pnCapitalVac, 0, 0, 0, 0, 0, 0, pnCapitalVac, 0, 0, IIf(pnMonedaPago = "2" And Mid(Mid(psEntTransferIFCod, 18, 10), 3, 1) = "1", 1, 0), pnImporte, 0, 0, 0, 0, 0, 0, pnCapitalVac, 0, 0, pdFecha, pdFecha, pdFecha, pdFecha, "0"
Else
    oDMov.GrabaCtaIFEstadistica psPersCod, CStr(Format(psIFTpo, "00")), psCuentaCod, "0", 0, lnMovNro, 1, pnCapitalVac, 0, 0, 0, 0, 0, 0, 0, 0, 0, IIf(pnMonedaPago = "2" And Mid(Mid(psEntTransferIFCod, 18, 10), 3, 1) = "1", 1, 0), pnImporte, 0, 0, 0, 0, 0, 0, pnCapitalVac, 0, 0, pdFecha, pdFecha, pdFecha, pdFecha, "0"
End If
If pbDatosContables = 1 Then
oDMov.ActualizaSaldoMovimiento psMovNro, "+"
End If
oDMov.CommitTrans
lbTrans = False
GrabaRegPagareAdeudado = 0
Set oDMov = Nothing
Set oContF = Nothing

Exit Function
GrabaRegPagareAdeudadoErr:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaAperturaCtaIF", lsMsgErr & " Error Graba Apertura de Ctas IF"
End Function

Public Function GrabaPagoCuotaAdeudadosLote(ByVal psMovNro As String, ByVal psOpecod As String, ByVal pdFecha As Date, ByVal psMovDesc As String, _
       ByVal rsAdeudados As ADODB.Recordset, _
       ByVal pdFecsis As Date, ByVal psEntTransferIFCod As String, _
       ByVal prsAsiento As ADODB.Recordset, ByVal pnTasaVac As Double, ByVal psCtaConcesional As String, _
       Optional psCtaOrdenD As String = "", Optional psCtaOrdenH As String = "") As Integer

Dim oContF     As NContFunciones
Dim oDMov      As DMov
Dim oOpe       As DOperacion
Dim oIF        As DInstFinanc

Dim lnMovItem  As Long
Dim lnMovOrden As Long
Dim lnMovNro   As Long
Dim lnMonto    As Currency
Dim lnMontoBco As Currency

Dim lsMovNro   As String
Dim lsSubCta   As String
Dim lbTrans    As Boolean
Dim lsCuenta   As String
Dim lsCtaCapital As String
Dim lsFiltro   As String

Dim lnCapital  As Currency
Dim lnInteres  As Currency
Dim lbExisteFiltro As Boolean
Dim nMontoConcesional As Currency
Dim nSaldoLP          As Currency
Dim lsMovNroLista     As String
Dim psCtaAdeudCod     As String

On Error GoTo GrabaPagoCuotaAdeudadosErr

Set oDMov = New DMov
Set oContF = New NContFunciones
Set oOpe = New DOperacion
Set oIF = New DInstFinanc


oDMov.BeginTrans
Do While Not rsAdeudados.EOF
    If rsAdeudados!Ok = "1" Then
        lbTrans = True
        If Mid(psMovNro, 20, 2) = "99" Then
           psMovNro = Left(psMovNro, 12) & LTrim(Str(Val(Mid(psMovNro, 13, 2)) + 1)) & Mid(psMovNro, 15, 5) & "00" & Mid(psMovNro, 22, 4)
        Else
           psMovNro = oDMov.GeneraMovNro(, , , psMovNro)
        End If
        psMovDesc = "Pago de Adeudados " & rsAdeudados!Entidad & " " & rsAdeudados!Adeudado
        oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable
        lnMovNro = oDMov.GetnMovNro(psMovNro)
        oDMov.InsertaMovCont lnMovNro, rsAdeudados![Total Pago], "0", "0"

        'Guardamos la cuenta en el debe
        lnMovItem = 0
        lnMovOrden = 0 'Se usara en este Bucle como Contador de Registro
        lnCapital = 0
        lnMontoBco = 0
        psCtaAdeudCod = rsAdeudados!objeto
        If Not prsAsiento Is Nothing Then
            If prsAsiento.State = adStateOpen Then
                If Not prsAsiento.EOF And Not prsAsiento.BOF Then
                    Do While Not prsAsiento.EOF
                        lnMovOrden = lnMovOrden + 1
                        lnMonto = 0
                        Select Case prsAsiento!Pos
                        Case "0"
                           lnMonto = rsAdeudados!Capital
                           lnCapital = rsAdeudados!Capital
                           lsCtaCapital = prsAsiento!Cuenta & lsSubCta
                        Case "1"
                           lnMonto = rsAdeudados![Int.Cal]   '10  y 11
                           lnInteres = rsAdeudados![Int.Cal]
                        Case "2"
                           lnMonto = rsAdeudados![Int.Prov.]   '8 y 9
                        Case "3"  'Comision
                           lnMonto = rsAdeudados![Comision]   '16
                        End Select
                        lnMontoBco = lnMontoBco + lnMonto

                        If lnMonto > 0 Then
                            lnMovItem = lnMovItem + 1
                            lsSubCta = ""
                            lsSubCta = oContF.GetFiltroObjetos(ObjEntidadesFinancieras, prsAsiento!Cuenta, psCtaAdeudCod, False, lbExisteFiltro)
                            If lsSubCta = "" And Not lbExisteFiltro Then
                                lsSubCta = oIF.GetIFSubCuenta(Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2))
                            End If
                            oDMov.InsertaMovCta lnMovNro, lnMovItem, prsAsiento!Cuenta & lsSubCta, lnMonto
                            oDMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
                            oDMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10)
                            Select Case lnMovOrden
                                Case 1
                                    lnCapital = rsAdeudados!Capital
                                    lsCtaCapital = prsAsiento!Cuenta & lsSubCta
                                Case 2
                                    lnInteres = prsAsiento![Monto VAC]
                            End Select
                        End If
                        prsAsiento.MoveNext
                    Loop
                    prsAsiento.MoveFirst
                Else
                    Err.Raise 50001, "nCajaGeneral:GrabaPagoCuotaAdeudadosLote", "Operacion no posee datos Contables necesarios"
                End If
            Else
                Err.Raise 50001, "nCajaGeneral:GrabaPagoCuotaAdeudadosLote", "Operacion no posee datos Contables necesarios"
            End If
        Else
            Err.Raise 50001, "nCajaGeneral:GrabaPagoCuotaAdeudadosLote", "Operacion no posee datos Contables necesarios"
        End If
        
        oDMov.ActualizaAdeudadosPagoCuota Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10), Format(pdFecha, gsFormatoFecha), lnCapital, lnInteres, rsAdeudados![Cuot], psMovNro, False
        If pnTasaVac > 0 And Mid(psOpecod, 3, 1) = "1" Then
            oDMov.InsertaMovTpoCambio lnMovNro, pnTasaVac
        End If
        
        'Transferencia
        If psEntTransferIFCod <> "" And lnMontoBco > 0 Then
            lnMovItem = lnMovItem + 1: lnMovOrden = 0
            lsCuenta = oOpe.EmiteOpeCta(psOpecod, "H", "1", psEntTransferIFCod, Left(psEntTransferIFCod, 2))
            oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta, lnMontoBco * -1
            lnMovOrden = lnMovOrden + 1
            oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
            oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psEntTransferIFCod, 4, 13), Mid(psEntTransferIFCod, 1, 2), Mid(psEntTransferIFCod, 18, 10)
        End If
        
        If psCtaOrdenD <> "" And lnCapital <> 0 Then 'Cuentas de Orden - Pepe
            lnMovItem = lnMovItem + 1
            oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaOrdenD, lnCapital
            lnMovItem = lnMovItem + 1
            oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaOrdenH, lnCapital * -1
        End If
        
        'Reclasificación Corto y Largo Plazo
        If rsAdeudados![nSaldoCapLP] <> 0 And lnCapital <> 0 Then
            oConect.ConexionActiva = oDMov.GetConexion
            nSaldoLP = CalculaSaldoAdeudadoLargoPlazo(pdFecha + 1, Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10))
            If rsAdeudados![nSaldoCapLP] <> nSaldoLP Then
                'Corto Plazo
                lnMovItem = lnMovItem + 1
                oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCtaCapital, nSaldoLP - rsAdeudados![nSaldoCapLP]
                oDMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
                oDMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10)
                
                'Largo Plazo
                lnMovItem = lnMovItem + 1
                oDMov.InsertaMovCta lnMovNro, lnMovItem, "26" + Mid(lsCtaCapital, 3) + "01", rsAdeudados![nSaldoCapLP] - nSaldoLP
                oDMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
                oDMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10)
                oDMov.ActualizaAdeudadosSaldoCap Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10), , nSaldoLP
            End If
        End If
        
        'Pago de Cuota Concesional
        'Adeudado Mi Vivienda
        nMontoConcesional = oDMov.CancelaCuotaConcesional(Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10), Format(pdFecha, gsFormatoFecha), gCGTipoCuotCalIFNoConcesional, psMovNro)
        If nMontoConcesional <> 0 And lnCapital <> 0 Then
            'Saldo Capital Adeudado
            lnMovItem = lnMovItem + 1
            oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCtaCapital, nMontoConcesional
            oDMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
            oDMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10)
        
            'Cuenta que revierte Pendiente de Creditos
            lnMovItem = lnMovItem + 1
            oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaConcesional, nMontoConcesional * -1
            oDMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
            oDMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10)
            oDMov.ActualizaAdeudadosSaldoCap Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10), nMontoConcesional
        End If
        
        lsMovNroLista = lsMovNroLista & "'" & psMovNro & "',"
        If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
            oDMov.GeneraMovME lnMovNro, psMovNro
        End If
    End If
    rsAdeudados.MoveNext
Loop
oDMov.ActualizaSaldoMovimiento Left(psMovNro, Len(psMovNro) - 1), "+"
oDMov.CommitTrans
lbTrans = False

Set oDMov = Nothing
Set oContF = Nothing

Exit Function
GrabaPagoCuotaAdeudadosErr:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaPagoCuotaAdeudados", lsMsgErr & " Error Graba Pago Adeudados"
End Function

Public Function GrabaPagoCuotaAdeudados(ByVal psMovNro As String, ByVal psOpecod As String, ByVal pdFecha As Date, ByVal psMovDesc As String, pnImporte As Currency, _
       ByVal psCtaAdeudCod As String, pnNroCuota As Integer, ByVal rsBillete As ADODB.Recordset, ByVal rsMoneda As ADODB.Recordset, _
       ByVal pdFecsis As Date, _
       ByVal psEntTransferIFCod As String, pnEntTransferIFImporte As Currency, _
       ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pdFechaDoc As Date, ByVal psNroVoucher As String, _
       ByVal prsCheque As ADODB.Recordset, ByVal prsOtro As ADODB.Recordset, ByVal prsObj As ADODB.Recordset, _
       ByVal prsAsiento As ADODB.Recordset, ByVal pbCancela As Boolean, ByVal pnTasaVac As Double, ByVal psCtaConcesional As String, Optional pnSaldoLP As Currency = 0, _
       Optional psCtaOrdenD As String = "", Optional psCtaOrdenH As String = "", Optional psCodLinCred As String) As Integer

Dim oContF     As NContFunciones
Dim oDMov      As DMov
Dim oOpe       As DOperacion
Dim oIF        As DInstFinanc

Dim lnMovItem  As Long
Dim lnMovOrden As Long
Dim lnMovNro   As Long
Dim lsMovNro   As String
Dim lsSubCta   As String
Dim lbTrans    As Boolean
Dim lsCuenta   As String
Dim lsCtaCapital As String
Dim lsFiltro   As String

Dim lnCapital  As Currency
Dim lnInteres  As Currency
Dim lbExisteFiltro As Boolean
Dim nMontoConcesional As Currency
Dim nSaldoLP          As Currency

On Error GoTo GrabaPagoCuotaAdeudadosErr

Set oDMov = New DMov
Set oContF = New NContFunciones
Set oOpe = New DOperacion
Set oIF = New DInstFinanc

GrabaPagoCuotaAdeudados = 1

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"

'Guardamos la cuenta en el debe
lnMovItem = 0
lnMovOrden = 0 'Se usara en este Bucle como Contador de Registro
lnCapital = 0
If Not prsAsiento Is Nothing Then
    If prsAsiento.State = adStateOpen Then
        If Not prsAsiento.EOF And Not prsAsiento.BOF Then
            Do While Not prsAsiento.EOF
                lnMovOrden = lnMovOrden + 1
                lsSubCta = ""
                lsSubCta = oContF.GetFiltroObjetos(ObjEntidadesFinancieras, prsAsiento!Cuenta, psCtaAdeudCod, False, lbExisteFiltro)
                If lsSubCta = "" And Not lbExisteFiltro Then
                   lsSubCta = oIF.GetIFSubCuenta(Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2))
                End If
                If prsAsiento!Monto > 0 Then
                    lnMovItem = lnMovItem + 1
                    oDMov.InsertaMovCta lnMovNro, lnMovItem, prsAsiento!Cuenta & lsSubCta, prsAsiento!Monto
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
                    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10)
                End If
                    Select Case lnMovOrden
                        Case 1
                            lnCapital = prsAsiento![Monto VAC]
                            lsCtaCapital = prsAsiento!Cuenta & lsSubCta
                        Case 2
                            lnInteres = prsAsiento![Monto VAC]
                    End Select
                prsAsiento.MoveNext
            Loop
        End If
    End If
End If

oDMov.ActualizaAdeudadosPagoCuota Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10), Format(pdFecha, gsFormatoFecha), lnCapital, lnInteres, pnNroCuota, psMovNro, pbCancela

'Actualiza Saldos de Linea de Credito
oDMov.ActualizaColocAdeudadoSaldo psCodLinCred, GetFechaMov(psMovNro, True), lnCapital * -1

If pnTasaVac > 0 And Mid(psOpecod, 3, 1) = "1" Then
    oDMov.InsertaMovTpoCambio lnMovNro, pnTasaVac
End If

'Guardamos el efectivo
lsCuenta = oOpe.EmiteOpeCta(psOpecod, "H", "0")
oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsBillete, lsCuenta, "H"
oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsMoneda, lsCuenta, "H"

'Transferencia
If psEntTransferIFCod <> "" And pnEntTransferIFImporte > 0 Then
    lnMovItem = lnMovItem + 1: lnMovOrden = 0
    lsCuenta = oOpe.EmiteOpeCta(psOpecod, "H", "1", psEntTransferIFCod, Left(psEntTransferIFCod, 2))
    oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta, pnEntTransferIFImporte * -1
    lnMovOrden = lnMovOrden + 1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psEntTransferIFCod, 4, 13), Mid(psEntTransferIFCod, 1, 2), Mid(psEntTransferIFCod, 18, 10)
End If

'Cheques recibidos - Cheque de Gerencia
If Not prsCheque Is Nothing Then
    If prsCheque.State = adStateOpen Then
        lsCuenta = oOpe.EmiteOpeCta(psOpecod, "H", "2")
        Do While Not prsCheque.EOF
            If prsCheque!Opc = "1" Then
                lnMovItem = lnMovItem + 1: lnMovOrden = 0
                
                lsFiltro = oContF.GetFiltroObjetos(ObjProductosCMACT, lsCuenta, prsCheque!Cuenta, False)
                lsFiltro = lsFiltro & oContF.GetFiltroObjetos(ObjCMACAgenciaArea, lsCuenta, prsCheque!cAreaCod & prsCheque!cAgeCod, False)
                oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta & lsFiltro, prsCheque!Importe * -1
                
                lnMovOrden = lnMovOrden + 1
                oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, prsCheque!Cuenta
                lnMovOrden = lnMovOrden + 1
                oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
                oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, prsCheque!cAgeCod, prsCheque!cAreaCod
                
                oDMov.ActualizaCheque TpoDocCheque, prsCheque!NroCheque, prsCheque!cPerscod, prsCheque!cIFTpo, , , , , , gCGEstadosChqDepositado
                oDMov.InsertaMovRef lnMovNro, prsCheque!nMovNro
            End If
            prsCheque.MoveNext
        Loop
    End If
End If

'Otras Formas de Pago
oDMov.GrabaMovOtrosCtaObj lnMovNro, lnMovItem, prsOtro, prsObj, "H"

If psCtaOrdenD <> "" Then 'Cuentas de Orden - Pepe
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaOrdenD, lnCapital
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaOrdenH, lnCapital * -1
End If

If psNroDoc <> "" Then
    oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, Format(pdFechaDoc, gsFormatoFecha)
End If
If psNroVoucher <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psNroVoucher, Format(pdFechaDoc, gsFormatoFecha)
End If

'Reclasificación Corto y Largo Plazo
If pnSaldoLP <> 0 Then
    oConect.ConexionActiva = oDMov.GetConexion
    nSaldoLP = CalculaSaldoAdeudadoLargoPlazo(pdFecha + 1, Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10))
    If pnSaldoLP <> nSaldoLP Then
        'Corto Plazo
        lnMovItem = lnMovItem + 1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCtaCapital, Round((nSaldoLP - pnSaldoLP) * IIf(pnTasaVac > 0, CDbl(Format(pnTasaVac, "#,##0.00###")), 1), 2)
        oDMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10)
        
        'Largo Plazo
        lnMovItem = lnMovItem + 1
        'oDMov.InsertaMovCta lnMovNro, lnMovItem, "26" + Mid(lsCtaCapital, 3) + "01", Round((pnSaldoLP - nSaldoLP) * IIf(pnTasaVac > 0, CDbl(Format(pnTasaVac, "#,##0.00###")), 1), 2)
        
        'Cambio
        If Mid(psOpecod, 3, 1) = "1" Then
            oDMov.InsertaMovCta lnMovNro, lnMovItem, "26" + Mid(lsCtaCapital, 3), Round((pnSaldoLP - nSaldoLP) * IIf(pnTasaVac > 0, CDbl(Format(pnTasaVac, "#,##0.00###")), 1), 2)
        Else
            oDMov.InsertaMovCta lnMovNro, lnMovItem, "26" + Mid(lsCtaCapital, 3) + "01", Round((pnSaldoLP - nSaldoLP) * IIf(pnTasaVac > 0, CDbl(Format(pnTasaVac, "#,##0.00###")), 1), 2)
        End If
        'Fin
        
        oDMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10)
        oDMov.ActualizaAdeudadosSaldoCap Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10), , nSaldoLP
    End If
End If

'Pago de Cuota Concesional
'Adeudado Mi Vivienda
nMontoConcesional = oDMov.CancelaCuotaConcesional(Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10), Format(pdFecha, gsFormatoFecha), gCGTipoCuotCalIFNoConcesional, psMovNro)
If nMontoConcesional <> 0 And lnCapital <> 0 Then
    'Saldo Capital Adeudado
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCtaCapital, nMontoConcesional
    oDMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10)

    'Cuenta que revierte Pendiente de Creditos
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaConcesional, nMontoConcesional * -1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10)
    oDMov.ActualizaAdeudadosSaldoCap Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10), nMontoConcesional
End If


If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"

oDMov.CommitTrans
lbTrans = False
GrabaPagoCuotaAdeudados = 0
Set oDMov = Nothing
Set oContF = Nothing

Exit Function
GrabaPagoCuotaAdeudadosErr:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaPagoCuotaAdeudadosLote", lsMsgErr & " Error Graba Pago Adeudados"
End Function

Public Function GrabaAperturaRegCheque(ByVal psMovNro As String, ByVal psOpecod As String, _
                                      ByVal psMovDesc As String, _
                                      ByVal psCtaDebe As String, ByVal psObjCodProd As String, ByVal psAreaCod As String, ByVal psAgeCod As String, _
                                      ByVal psCtaHaber As String, ByVal rsObjHaber As ADODB.Recordset, _
                                      ByVal psNumCheque As String, ByVal psPersCodIf As String, ByVal psTipoIF As String, ByVal pnPlaza As ChequePlaza, _
                                      ByVal psCtaCheque As String, ByVal pnImporte As Currency, ByVal pdFechaReg As Date, _
                                      ByVal pdFechaVal As Date, ByVal pnMonedaCheque As Moneda, _
                                      ByVal pbCreaSubCta As Boolean, ByVal psSubCtaIFCod As String, ByVal psSubCtaIFDesc As String, _
                                      ByVal psPersCod As String, ByVal psIFTpo As String, _
                                      ByVal psCtaIFCod As String, ByVal psCtaIFDesc As String, _
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
oMov.BeginTrans
lbTrans = True
oMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabNoContable
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovCont lnMovNro, pnImporte, 0, ""

'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1
oMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe, pnImporte
lnMovOrden = lnMovOrden + 1

If pbCreaSubCta Then
    oMov.InsertaCuentaIF psPersCod, psIFTpo, psSubCtaIFCod, psSubCtaIFDesc, pdCtaIFAper, Format(pdCtaIFAper, gsFormatoFecha), Format(pdCtaIFAper, gsFormatoFecha), 0, CDate(psCtaIFVenc), gEstadoCtaIFActiva, psMovNro
End If
oMov.InsertaCuentaIF psPersCod, psIFTpo, psCtaIFCod, psCtaIFDesc, pdCtaIFAper, psCtaIFVenc, pdCtaIFAper, _
                        pnCtaIFPlazo, CDate(psCtaIFVenc), gEstadoCtaIFRegistrada, psMovNro

oMov.InsertaCuentaIFInteres psPersCod, psIFTpo, psCtaIFCod, pdCtaIFAper, pnPeriodo, pnInteres, psMovNro

oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod

'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
'oDMov.InsertaMovCtaIF lnMovNro, psPersCod, psIFTpo, psCtaIFCod, CGCtaIFConCapital, pnImporte
'oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                            psPersCod , psIFTpo, psCtaIfCod

lnMovItem = lnMovItem + 1: lnMovOrden = 1
oMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber, pnImporte * -1
oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, psObjCodProd
lnMovOrden = lnMovOrden + 1
oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
oMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, psAgeCod, psAreaCod

If psNumCheque <> "" Then
    oMov.InsertaMovDoc lnMovNro, TpoDocCheque, psNumCheque, Format(pdFechaReg, gsFormatoFecha)
End If
'grabacion dentro de la tabla de cheques
oMov.InsertaCheque TpoDocCheque, psNumCheque, psPersCodIf, psTipoIF, pnPlaza, psCtaCheque, pnImporte, pdFechaReg, _
                    pdFechaVal, pnEstadoCG, pnConfirmaCaja, pnMonedaCheque, , psAreaCodChq, psAgeCodChq

oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIf, psTipoIF, pdFechaReg, gChqEstRegistrado, psMovNro, psCtaCheque

If psNroDocAper <> "" Then
    oMov.InsertaMovDoc lnMovNro, pnTpoDocAper, psNroDocAper, Format(pdFechaDocApera, gsFormatoFecha)
End If
If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
    oMov.GeneraMovME lnMovNro, psMovNro
End If
oMov.ActualizaSaldoMovimiento psMovNro, "+"
oMov.CommitTrans
lbTrans = False

GrabaAperturaRegCheque = 0
Exit Function
GrabaAperturaRegChequeErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub GeneraCuentaBancos(psCuenta As String, psSubCuenta As String, psPersCod As String, psIFTpo As String, psCtaIFCod As String, psOpecod As String, Optional psDH As String = "D")
Dim lsMismaSubCta As Boolean
'Verificamos si Existe en CtaIFFiltro - segun Cuenta definidas en Operacion
sql = "SELECT Distinct cif.cCtaContCod, cif.cCtaIFSubCta FROM CtaIFFiltro cif JOIN OpeCta oc ON cif.cCtaContCod = oc.cCtaContCod " _
    & "WHERE cif.cPersCod = '" & psPersCod & "' and cif.cCtaIFCod LIKE '" & Left(psCtaIFCod, 2) & "%' and oc.cOpeCod = '" & psOpecod & "' and cOpeCtaDH = '" & psDH & "' and cOpeCtaOpc = '1' "
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF Then
    psCuenta = rs!cCtaContCod
    If Val(Left(psCtaIFCod, 2)) = gTpoCtaIFCtaPF Or Val(Left(psCtaIFCod, 2)) = gTpoCtaIFCtaAdeud Then
        psSubCuenta = rs!cCtaIFSubCta
    Else
        If rs.RecordCount = 1 Then
            'Si solo hay uno => Se crearia Hijos. el Q existe seria 01 y este nuevo 02
            psSubCuenta = rs!cCtaIFSubCta '& "02"
            'Rs.MoveLast
            'psSubCuenta = Format(Val(Rs!cCtaIFSubCta) + 1, String(Len(Rs!cCtaIFSubCta), "0"))
        Else
            rs.MoveLast
            'psSubCuenta = Format(Val(Rs!cCtaIFSubCta) + 1, String(Len(Rs!cCtaIFSubCta), "0"))
            psSubCuenta = rs!cCtaIFSubCta
        End If
    End If
Else
    
    'Verificamos si IF ya tiene Cuentas de Este Tipo
    sql = "SELECT Distinct cif.cCtaContCod FROM CtaIFFiltro cif JOIN OpeCta oc ON cif.cCtaContCod = oc.cCtaContCod " _
        & "WHERE cIFTpo = '" & psIFTpo & "' and cCtaIFCod LIKE '" & Left(psCtaIFCod, 2) & "%' and oc.cOpeCod = '" & psOpecod & "'"
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF Then
        If rs.RecordCount = 1 Then
            psCuenta = rs!cCtaContCod
        End If
    End If 'BANCO NO TIENE NINGUNA CUENTA
    If psSubCuenta = "" Then
        sql = "SELECT cSubCtaContCod FROM InstitucionFinanc WHERE cPersCod = '" & psPersCod & "'"
        Set rs = oConect.CargaRecordSet(sql)
        If Not rs.EOF Then
            psSubCuenta = rs!cSubCtaContCod & IIf((Left(psCuenta, 4) = "1113" Or Left(psCuenta, 4) = "1123") And Left(psCtaIFCod, 2) = "03", "03", Left(psCtaIFCod, 2))
        End If
        
    End If
End If
'Verificamos Si existe la Cuenta Contable a Generar
sql = "SELECT cCtaContCod FROM CtaCont WHERE cCtaContCod = '" & psCuenta & psSubCuenta & "'"
Set rs = oConect.CargaRecordSet(sql)
If rs.EOF Then
    MsgBox "Cuenta Generada [" & psCuenta & psSubCuenta & "] no Existe. Consultar con Contabilidad o Sistemas"
End If

rs.Close

End Sub
'EJVG20120802 ***
Public Sub GeneraCuentaBancos2(psCuenta As String, psSubCuenta As String, psPersCod As String, psIFTpo As String, psCtaIFCod As String, psOpecod As String, Optional psDH As String = "D")
    Dim rsCtaBCRP As ADODB.Recordset 'PASI20151029
    
    psCuenta = "11" & Mid(psOpecod, 3, 1) & "3" & Format(psIFTpo, "00")
    sql = "SELECT cSubCtaContCod FROM InstitucionFinanc WHERE cPersCod = '" & psPersCod & "'"
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF Then
        psSubCuenta = rs!cSubCtaContCod & Left(psCtaIFCod, 2)
    End If
    
    'BCRP
    If psPersCod = "1090100822183" Then
        Select Case Left(psCtaIFCod, 2)
            'Modificado xPASI20151029 ******************
            'Case "03" 'DPF
            '    psCuenta = "11" & Mid(psOpecod, 3, 1) & "2"
            '    psSubCuenta = "0901"
            'Case "04" 'OverNight
            '    psCuenta = "11" & Mid(psOpecod, 3, 1) & "2"
            '    psSubCuenta = "0401"
            
            Case "03" 'DPF
                Set rsCtaBCRP = ObtieneCtaxConfirmBCRP("03")
                psCuenta = Replace(rsCtaBCRP!cCta, "M", Mid(psOpecod, 3, 1))
                psSubCuenta = rsCtaBCRP!cSubCta
            Case "04" 'OverNight
                Set rsCtaBCRP = ObtieneCtaxConfirmBCRP("04")
                psCuenta = Replace(rsCtaBCRP!cCta, "M", Mid(psOpecod, 3, 1))
                psSubCuenta = rsCtaBCRP!cSubCta
            'Pasi end **********************************
        End Select
    End If
    
    'Verificamos Si existe la Cuenta Contable a Generar
    sql = "SELECT cCtaContCod FROM CtaCont WHERE cCtaContCod = '" & psCuenta & psSubCuenta & "'"
    Set rs = oConect.CargaRecordSet(sql)
    If rs.EOF Then
        MsgBox "Cuenta Generada [" & psCuenta & psSubCuenta & "] no Existe. Consultar con Contabilidad", vbInformation, "Aviso"
    End If
    Set rs = Nothing
End Sub
'END EJVG *******
Public Sub GeneraCuentaAdeudos(psCuenta As String, psSubCuenta As String, psPersCod As String, psIFTpo As String, psCtaIFCod As String, psOpecod As String, Optional psDH As String = "D")
Dim lsMismaSubCta As Boolean
'Verificamos si Existe en CtaIFFiltro - segun Cuenta definidas en Operacion
sql = "SELECT Distinct cif.cCtaContCod, cif.cCtaIFSubCta FROM CtaIFFiltro cif JOIN OpeCta oc ON cif.cCtaContCod = oc.cCtaContCod " _
    & "WHERE cif.cPersCod = '" & psPersCod & "' and cif.cCtaIFCod LIKE '" & Left(psCtaIFCod, 2) & "%' and oc.cOpeCod = '" & psOpecod & "' and cOpeCtaDH = '" & psDH & "' and cOpeCtaOpc = '1' "
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF Then
    psCuenta = rs!cCtaContCod
    psSubCuenta = rs!cCtaIFSubCta
Else
    
    'Verificamos si IF ya tiene Cuentas de Este Tipo
    sql = "SELECT Distinct cif.cCtaContCod FROM CtaIFFiltro cif JOIN OpeCta oc ON cif.cCtaContCod = oc.cCtaContCod " _
        & "WHERE cIFTpo = '" & psIFTpo & "' and cCtaIFCod LIKE '" & Left(psCtaIFCod, 2) & "%' and oc.cOpeCod = '" & psOpecod & "'"
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF Then
        If rs.RecordCount = 1 Then
            psCuenta = rs!cCtaContCod
        End If
    End If 'BANCO NO TIENE NINGUNA CUENTA
    If psSubCuenta = "" Then
        sql = "SELECT cSubCtaContCod FROM InstitucionFinanc WHERE cPersCod = '" & psPersCod & "'"
        Set rs = oConect.CargaRecordSet(sql)
        If Not rs.EOF Then
            psSubCuenta = rs!cSubCtaContCod & IIf(Left(psCuenta, 4) = "1113" Or Left(psCuenta, 4) = "1123" And Left(psCtaIFCod, 2) = "03", "02", Left(psCtaIFCod, 2))
        End If
        
    End If
End If
'Verificamos Si existe la Cuenta Contable a Generar
sql = "SELECT cCtaContCod FROM CtaCont WHERE cCtaContCod = '" & psCuenta & psSubCuenta & "'"
Set rs = oConect.CargaRecordSet(sql)
If rs.EOF Then
    MsgBox "Cuenta Generada [" & psCuenta & psSubCuenta & "] no Existe. Consultar con Contabilidad o Sistemas"
End If

rs.Close

End Sub
Public Function GetApertSinConf(ByVal psOpecod As String, ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset
Dim lsFiltroAge As String
Set rs = New ADODB.Recordset

sql = "Select   CONVERT(CHAR(10),CONVERT(DATETIME, SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHA," _
    & "         P.CPERSNOMBRE AS BANCO, SC.cCtaIFDesc + ' ' + C.cCtaIFDesc AS CUENTA, " _
    & "         CONVERT(CHAR(10),C.dCtaIFAper,103) AS APERTURA, ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)  AS IMPORTE,  M.NMOVNRO, M.COPECOD, C.CCTAIFCOD, M.CMOVDESC, C.cIFTpo, C.cPersCod, ISNull(C.nTpoPF,0) nTpoPF" _
    & " From    Mov M " _
    & "         JOIN MovCta MC ON MC.nMovNro =M.nMovNro " _
    & "         LEFT JOIN MOVME ME ON ME.NMOVNRO=MC.NMOVNRO and me.nMovItem = mc.nMovItem " _
    & "         JOIN MOVOBJ MO on MO.nMovNro = MC.nMovNro and MO.nmovitem =MC.NMOVITEM " _
    & "         JOIN MOVOBJIF MIF ON MIF.nMovNro = MO.nMovNro and MIF.nMovItem= MO.NMOVITEM AND MO.NMOVOBJORDEN = MIF.NMOVOBJORDEN " _
    & "         JOIN CTAIF C ON C.cPersCod =MIF.cPersCod AND C.cIFTpo = MIF.cIFTpo AND C.cCtaIfCod = MIF.cCtaIfCod " _
    & "         JOIN CTAIF SC ON SC.cPersCod =MIF.cPersCod AND SC.cIFTpo = MIF.cIFTpo AND SUBSTRING(MIF.cCtaIfCod,1,LEN(SC.cCtaIfCod)) = SC.cCtaIfCod AND LEN(SC.cCtaIfCod)<=5 " _
    & "         JOIN PERSONA P ON P.CPERSCOD = C.CPERSCOD " _
    & "         JOIN OPECTA OC ON OC.cCtaContCod= MC.cCtaContCod " _
    & "         JOIN OPEOBJ OO ON OO.COPECOD = OC.COPECOD  " _
    & " Where   OC.COPECOD ='" & psOpecod & "' AND cOpeCtaDH='H' AND MC.NMOVIMPORTE>0 AND M.NMOVFLAG =" & gMovFlagVigente & "" _
    & "         AND C.cIFTpo + C.CCTAIFCOD LIKE cOpeObjFiltro  " _
    & "         AND SUBSTRING(M.cMovnro,1,8 ) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "'  " _
    & "         AND c.cCtaIFEstado = " & gEstadoCtaIFRegistrada
    
Set rs = oConect.CargaRecordSet(sql)
Set GetApertSinConf = rs

End Function

''' Nuevo Pepe Trujillo
'Public Function GetPagareSinConf(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, Optional psEstado As CGEstadoCtaIF = gEstadoCtaIFRegistrada) As ADODB.Recordset
'Dim lsFiltroAge As String
'Set Rs = New ADODB.Recordset
'
'sql = "Select   CONVERT(CHAR(10),CONVERT(DATETIME, SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHA," _
'    & "         P.CPERSNOMBRE AS BANCO, SC.cCtaIFDesc + ' ' + C.cCtaIFDesc AS CUENTA, " _
'    & "         CONVERT(CHAR(10),C.dCtaIFAper,103) AS APERTURA, ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE))  AS IMPORTE,  M.NMOVNRO, M.COPECOD, C.CCTAIFCOD, M.CMOVDESC, C.cIFTpo, C.cPersCod, m.cMovNro  " _
'    & " From    Mov M " _
'    & "         JOIN MovCta MC ON MC.nMovNro =M.nMovNro " _
'    & "         LEFT JOIN MOVME ME ON ME.NMOVNRO=MC.NMOVNRO and me.nMovItem = mc.nMovItem " _
'    & "         JOIN MOVOBJ MO on MO.nMovNro = MC.nMovNro and MO.nmovitem =MC.NMOVITEM " _
'    & "         JOIN MOVOBJIF MIF ON MIF.nMovNro = MO.nMovNro and MIF.nMovItem= MO.NMOVITEM AND MO.NMOVOBJORDEN = MIF.NMOVOBJORDEN " _
'    & "         JOIN CTAIF C ON C.cPersCod =MIF.cPersCod AND C.cIFTpo = MIF.cIFTpo AND C.cCtaIfCod = MIF.cCtaIfCod " _
'    & "         JOIN CTAIF SC ON SC.cPersCod =MIF.cPersCod AND SC.cIFTpo = MIF.cIFTpo AND SUBSTRING(MIF.cCtaIfCod,1,LEN(SC.cCtaIfCod)) = SC.cCtaIfCod AND LEN(SC.cCtaIfCod)<=5 " _
'    & "         JOIN PERSONA P ON P.CPERSCOD = C.CPERSCOD " _
'    & "         JOIN OPECTA OC ON OC.cCtaContCod= MC.cCtaContCod " _
'    & "         JOIN OPEOBJ OO ON OO.COPECOD = OC.COPECOD  " _
'    & " Where   OC.COPECOD ='" & psOpeCod & "' AND cOpeCtaDH='D' AND MC.NMOVIMPORTE" & IIf(psEstado = gEstadoCtaIFRegistrada, "<", ">") & "0 AND M.NMOVFLAG =" & gMovFlagVigente & "" _
'    & "         AND C.cIFTpo + C.CCTAIFCOD LIKE cOpeObjFiltro  " _
'    & "         AND SUBSTRING(M.cMovnro,1,8 ) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "'  " _
'    & "         AND c.cCtaIFEstado = " & psEstado
'
'Set Rs = oConect.CargaRecordSet(sql)
'Set GetPagareSinConf = Rs
'End Function


Public Function GetPagareSinConf(ByVal psOpecod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, Optional psEstado As CGEstadoCtaIF = gEstadoCtaIFRegistrada) As ADODB.Recordset
Dim lsFiltroAge As String
Set rs = New ADODB.Recordset
'Dim oCon As New DConecta
Dim sSql As String
    
'    oCon.AbreConexion
    sSql = "Select   CONVERT(CHAR(10),CONVERT(DATETIME, SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHA,"
    sSql = sSql & "         P.CPERSNOMBRE AS BANCO, SC.cCtaIFDesc + ' ' + C.cCtaIFDesc AS CUENTA, "
    sSql = sSql & "         CONVERT(CHAR(10),C.dCtaIFAper,103) AS APERTURA, ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE))  AS IMPORTE,  M.NMOVNRO, M.COPECOD, C.CCTAIFCOD, M.CMOVDESC, C.cIFTpo, C.cPersCod, m.cMovNro, "
    sSql = sSql & "         CIA.nMontoPrestado, CIA.nMontoPrestadoReal, CIA.cMonedaPago "
    sSql = sSql & " From    Mov M "
    sSql = sSql & "         JOIN MovCta MC ON MC.nMovNro =M.nMovNro "
    sSql = sSql & "         LEFT JOIN MOVME ME ON ME.NMOVNRO=MC.NMOVNRO and me.nMovItem = mc.nMovItem "
    sSql = sSql & "         JOIN MOVOBJ MO on MO.nMovNro = MC.nMovNro and MO.nmovitem =MC.NMOVITEM "
    sSql = sSql & "         JOIN MOVOBJIF MIF ON MIF.nMovNro = MO.nMovNro and MIF.nMovItem= MO.NMOVITEM AND MO.NMOVOBJORDEN = MIF.NMOVOBJORDEN "
    sSql = sSql & "         JOIN CTAIF C ON C.cPersCod =MIF.cPersCod AND C.cIFTpo = MIF.cIFTpo AND C.cCtaIfCod = MIF.cCtaIfCod "
    sSql = sSql & "         JOIN CTAIF SC ON SC.cPersCod =MIF.cPersCod AND SC.cIFTpo = MIF.cIFTpo AND SUBSTRING(MIF.cCtaIfCod,1,LEN(SC.cCtaIfCod)) = SC.cCtaIfCod AND LEN(SC.cCtaIfCod)<=5 "
    sSql = sSql & "         JOIN CTAIFADEUDADOS CIA ON C.cPersCod = CIA.cPersCod And C.cIFTpo=CIA.cIFTpo And C.cCtaIFCod=CIA.cCtaIFCod "
    sSql = sSql & "         JOIN PERSONA P ON P.CPERSCOD = C.CPERSCOD "
    sSql = sSql & "         JOIN OPECTA OC ON OC.cCtaContCod= MC.cCtaContCod "
    sSql = sSql & "         JOIN OPEOBJ OO ON OO.COPECOD = OC.COPECOD  "
    sSql = sSql & " Where   OC.COPECOD ='" & psOpecod & "' AND cOpeCtaDH='D' AND MC.NMOVIMPORTE" & IIf(psEstado = gEstadoCtaIFRegistrada, "<", ">") & "0 AND M.NMOVFLAG =" & gMovFlagVigente & ""
    sSql = sSql & "         AND C.cIFTpo + C.CCTAIFCOD LIKE cOpeObjFiltro  "
    sSql = sSql & "         AND SUBSTRING(M.cMovnro,1,8 ) >='" & Format(pdDesde, "yyyymmdd") & "' AND substring(M.cMovNro, 1,8) <='" & Format(pdHasta, "yyyymmdd") & "'  "
    sSql = sSql & "         AND c.cCtaIFEstado = " & psEstado
    
Set rs = oConect.CargaRecordSet(sSql)

'Set GetListaAdeudosPorConfirmar = rs
Set GetPagareSinConf = rs
End Function
Public Function GetPagareExt(ByVal psOpecod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, Optional psEstado As CGEstadoCtaIF = gEstadoCtaIFRegistrada) As ADODB.Recordset
Dim lsFiltroAge As String
Set rs = New ADODB.Recordset
'Dim oCon As New DConecta
Dim sSql As String
    
'    oCon.AbreConexion
    sSql = "Select   CONVERT(CHAR(10),CONVERT(DATETIME, SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHA,"
    sSql = sSql & "         P.CPERSNOMBRE AS BANCO, SC.cCtaIFDesc + ' ' + C.cCtaIFDesc AS CUENTA, "
    sSql = sSql & "         CONVERT(CHAR(10),C.dCtaIFAper,103) AS APERTURA, ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE))  AS IMPORTE,  M.NMOVNRO, M.COPECOD, C.CCTAIFCOD, M.CMOVDESC, C.cIFTpo, C.cPersCod, m.cMovNro, "
    sSql = sSql & "         CIA.nMontoPrestado, CIA.nMontoPrestadoReal, CIA.cMonedaPago "
    sSql = sSql & " From    Mov M "
    sSql = sSql & "         JOIN MovCta MC ON MC.nMovNro =M.nMovNro "
    sSql = sSql & "         LEFT JOIN MOVME ME ON ME.NMOVNRO=MC.NMOVNRO and me.nMovItem = mc.nMovItem "
    sSql = sSql & "         JOIN MOVOBJ MO on MO.nMovNro = MC.nMovNro and MO.nmovitem =MC.NMOVITEM "
    sSql = sSql & "         JOIN MOVOBJIF MIF ON MIF.nMovNro = MO.nMovNro and MIF.nMovItem= MO.NMOVITEM AND MO.NMOVOBJORDEN = MIF.NMOVOBJORDEN "
    sSql = sSql & "         JOIN CTAIF C ON C.cPersCod =MIF.cPersCod AND C.cIFTpo = MIF.cIFTpo AND C.cCtaIfCod = MIF.cCtaIfCod "
    sSql = sSql & "         JOIN CTAIF SC ON SC.cPersCod =MIF.cPersCod AND SC.cIFTpo = MIF.cIFTpo AND SUBSTRING(MIF.cCtaIfCod,1,LEN(SC.cCtaIfCod)) = SC.cCtaIfCod AND LEN(SC.cCtaIfCod)<=5 "
    sSql = sSql & "         JOIN CTAIFADEUDADOS CIA ON C.cPersCod = CIA.cPersCod And C.cIFTpo=CIA.cIFTpo And C.cCtaIFCod=CIA.cCtaIFCod "
    sSql = sSql & "         JOIN PERSONA P ON P.CPERSCOD = C.CPERSCOD "
    sSql = sSql & "         JOIN OPECTA OC ON OC.cCtaContCod= MC.cCtaContCod "
    sSql = sSql & "         JOIN OPEOBJ OO ON OO.COPECOD = OC.COPECOD  "
    sSql = sSql & " Where   OC.COPECOD ='" & psOpecod & "' AND cOpeCtaDH='H' AND MC.NMOVIMPORTE" & IIf(psEstado = gEstadoCtaIFRegistrada, "<", ">") & "0 AND M.NMOVFLAG =" & gMovFlagVigente & ""
    sSql = sSql & "         AND C.cIFTpo + C.CCTAIFCOD LIKE cOpeObjFiltro  "
    sSql = sSql & "         AND SUBSTRING(M.cMovnro,1,8 ) >='" & Format(pdDesde, "yyyymmdd") & "' AND substring(M.cMovNro, 1,8) <='" & Format(pdHasta, "yyyymmdd") & "'  "
    sSql = sSql & "         AND c.cCtaIFEstado = " & psEstado
    
Set rs = oConect.CargaRecordSet(sSql)

'Set GetListaAdeudosPorConfirmar = rs
Set GetPagareExt = rs
End Function




'' Nuevo Pepe Trujillo
Public Function CalculaSaldoAdeudadoLargoPlazo(pdFecha As Date, psPersCod As String, psIFTpo As String, psCtaIFCod As String, Optional psTpoCuota As CGTipoCuotCalIF = gCGTipoCuotCalIFCuota, Optional pCon As Object) As Currency
On Error GoTo CalculaSaldoAdeudadoLargoPlazoErr
CalculaSaldoAdeudadoLargoPlazo = 0
    sql = "SELECT ISNULL(SUM(nCapital),0) nCapital FROM CtaIFCalendario " _
        & "WHERE cPersCod = '" & psPersCod & "' and cIFTpo = '" & psIFTpo & "' and cCtaIFCod = '" & psCtaIFCod & "' and " _
        & "      datediff(d,'" & Format(pdFecha, gsFormatoFecha) & "',dVencimiento) > 365 and cEstado = '" & gTpoEstCuotaAdeudPend & "' " _
        & "      and bVigente=1 "
    'ALPA20130617*****************Se agregó bVigente=1
    If Not pCon Is Nothing Then
        oConect.ConexionActiva = pCon
        lbConexionPropia = False
    End If
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF Then
        CalculaSaldoAdeudadoLargoPlazo = rs!nCapital
    End If
    'RSClose Rs
    rs.Close
    Exit Function
CalculaSaldoAdeudadoLargoPlazoErr:
    Err.Raise vbObjectError + 100, "CalculaSaldoAdeudadoLargoPlazo", Err.Description & " Adeudados "
End Function

''' Nuevo Pepe Trujillo
'Public Function GrabaConfApertura(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
'                                ByVal psCtaDebe As String, ByVal psSubCtaDebe As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
'                                ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psCtaIFCod As String, _
'                                ByVal pnMovNroRef As Long, ByVal pdFecha As Date, Optional ByVal pnTasaVac As Double) As Integer
'Dim lnMovItem As Integer
'Dim lnMovOrden As Integer
'Dim oDMov As DMov
'Dim lsMovNro As String
'Dim lsSubCta As String
'Dim lbTrans As Boolean
'Dim lnMovNro As Long
'Dim oCtaIf As NCajaCtaIF
'Dim lnSaldo As Currency
'Dim ldFecha As Date
'Dim lnFactor As Integer
'Dim nImporteLP As Currency
'
'Set oCtaIf = New NCajaCtaIF
'Set oDMov = New DMov
'On Error GoTo GrabaConfAperturaErr
'nImporteLP = 0
'lnFactor = 1
'
'If Val(Left(psCtaIFCod, 2)) = gTpoCtaIFCtaAdeud Then
'   nImporteLP = CalculaSaldoAdeudadoLargoPlazo(GetFechaMov(psMovNro, True), psPersCod, psIFTpo, psCtaIFCod)
'   lnFactor = -1
'End If
'
'GrabaConfApertura = 1
'oDMov.BeginTrans
'lbTrans = True
'oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc
'lnMovNro = oDMov.GetnMovNro(psMovNro)
'ldFecha = CDate(GetFechaMov(psMovNro, True))
'
'oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
''guardamos la cuenta de Bancos en el debe : Primer Capital
'lnMovItem = 0: lnMovOrden = 0
'lnMovItem = lnMovItem + 1
'oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe & psSubCtaDebe, (pnImporte - Format(nImporteLP * IIf(pnTasaVac > 0, pnTasaVac, 1), "#,#0.00")) * lnFactor
'lnMovOrden = lnMovOrden + 1
'oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
'oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod
'
''Creamos la SubCuenta de la Inst.Financ en CtaIFFiltro
'oDMov.InsertaCuentaIFFiltro psPersCod, psIFTpo, psCtaIFCod, psCtaDebe, psSubCtaDebe
'
''ADEUDADOS Largo Plazo
'If nImporteLP > 0 Then
'    lnMovOrden = 0
'    lnMovItem = lnMovItem + 1
'    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), "26" & Mid(psCtaDebe, 3, 20) & psSubCtaDebe, Format(nImporteLP * lnFactor * IIf(pnTasaVac > 0, pnTasaVac, 1), "#,#0.00")
'    lnMovOrden = lnMovOrden + 1
'    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
'    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod
'    oDMov.ActualizaAdeudadosSaldoCap psPersCod, psIFTpo, psCtaIFCod, , nImporteLP
'End If
'
''Guardamos la Cuenta de la Pendiente que se elimina
'lnMovItem = lnMovItem + 1: lnMovOrden = 0
'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, (pnImporte * -1) * lnFactor
'lnMovOrden = lnMovOrden + 1
'oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
'oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod
'
''Pasamos a activo el estado de la cuenta aperturada
'oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIFCod, , , Format(ldFecha, gsFormatoFecha), Format(ldFecha, gsFormatoFecha), , , gEstadoCtaIFActiva
'oDMov.InsertaMovRef lnMovNro, pnMovNroRef
'
'If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
'    oDMov.GeneraMovME lnMovNro, psMovNro
'End If
'oDMov.ActualizaSaldoMovimiento psMovNro, "+"
'
'oDMov.CommitTrans
'lbTrans = False
'
'GrabaConfApertura = 0
'Set oDMov = Nothing
'
'Exit Function
'GrabaConfAperturaErr:
'    lsMsgErr = Err.Description
'    If lbTrans Then
'        oDMov.RollbackTrans
'        lbTrans = False
'    End If
'    Err.Raise vbObjectError + 100, "GrabaConfApertura", lsMsgErr & " Error en Grabación de Confirmación de Apertura "
'End Function


Public Function GrabaConfApertura(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
                                ByVal psCtaDebe As String, ByVal psSubCtaDebe As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psCtaIFCod As String, _
                                ByVal pnMovNroRef As Long, ByVal pdFecha As Date, Optional ByVal pnTasaVac As Double) As Integer
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim oCtaIf As NCajaCtaIF
Dim lnSaldo As Currency
Dim ldFecha As Date
Dim lnFactor As Integer
Dim nImporteLP As Currency
Dim lsCuenta As String
Dim lsCuentaInt As String 'EJVG20111012

Set oCtaIf = New NCajaCtaIF
Set oDMov = New DMov
On Error GoTo GrabaConfAperturaErr
nImporteLP = 0
lnFactor = 1

If Val(Left(psCtaIFCod, 2)) = gTpoCtaIFCtaAdeud Then
   nImporteLP = CalculaSaldoAdeudadoLargoPlazo(GetFechaMov(psMovNro, True), psPersCod, psIFTpo, psCtaIFCod)
   lnFactor = -1
End If

GrabaConfApertura = 1
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpecod, psMovDesc
lnMovNro = oDMov.GetnMovNro(psMovNro)
ldFecha = CDate(GetFechaMov(psMovNro, True))

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta de Bancos en el debe : Primer Capital
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe & psSubCtaDebe, (pnImporte - Format(nImporteLP * IIf(pnTasaVac > 0, pnTasaVac, 1), "#,#0.00")) * lnFactor
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod

'Creamos la SubCuenta de la Inst.Financ en CtaIFFiltro
oDMov.InsertaCuentaIFFiltro psPersCod, psIFTpo, psCtaIFCod, psCtaDebe, psSubCtaDebe
'Creamos la SubCuenta de la Inst.Financ en CtaIFFiltro - para intereses 5
'Gitu 08-09-2008
'If psOpeCod = "402515" Or psOpeCod = "401515" Then
    'lsCuenta = "51M10301"
    'lsCuentaInt = "11M80301" 'EJVG20111012
'Else
    'lsCuenta = "51M10303"
    'lsCuentaInt = "11M80303" 'EJVG20111012
'End If
'Fin gitu

lsCuenta = "51M103" & Format(psIFTpo, "00")
lsCuentaInt = "11M803" & Format(psIFTpo, "00")

oDMov.InsertaCuentaIFFiltro psPersCod, psIFTpo, psCtaIFCod, Replace(lsCuenta, "M", Mid(psCtaDebe, 3, 1)), psSubCtaDebe
'Creamos la SubCuenta de la Inst.Financ en CtaIFFiltro - para intereses 11M8
oDMov.InsertaCuentaIFFiltro psPersCod, psIFTpo, psCtaIFCod, Replace(lsCuentaInt, "M", Mid(psCtaDebe, 3, 1)), psSubCtaDebe

'ADEUDADOS Largo Plazo
If nImporteLP > 0 Then
    lnMovOrden = 0
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), "26" & Mid(psCtaDebe, 3, 20) & psSubCtaDebe, Format(nImporteLP * lnFactor * IIf(pnTasaVac > 0, pnTasaVac, 1), "#,#0.00")
    lnMovOrden = lnMovOrden + 1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod
    oDMov.ActualizaAdeudadosSaldoCap psPersCod, psIFTpo, psCtaIFCod, , nImporteLP
End If

'Guardamos la Cuenta de la Pendiente que se elimina
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, (pnImporte * -1) * lnFactor
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod

'Pasamos a activo el estado de la cuenta aperturada
oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIFCod, , , Format(ldFecha, gsFormatoFecha), Format(ldFecha, gsFormatoFecha), , , gEstadoCtaIFActiva
oDMov.InsertaMovRef lnMovNro, pnMovNroRef

If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"

oDMov.CommitTrans
lbTrans = False

GrabaConfApertura = 0
Set oDMov = Nothing

Exit Function
GrabaConfAperturaErr:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaConfApertura", lsMsgErr & " Error en Grabación de Confirmación de Apertura "
End Function

'EJVG20120802 ***
Public Function GrabaConfApertura2(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
                                ByVal psCtaDebe As String, ByVal psSubCtaDebe As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psCtaIFCod As String, _
                                ByVal pnMovNroRef As Long, ByVal pdFecha As Date, Optional ByVal pnTasaVac As Double) As Integer
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim oCtaIf As NCajaCtaIF
Dim lnSaldo As Currency
Dim ldFecha As Date
Dim lnFactor As Integer
Dim nImporteLP As Currency
Dim lsCuenta As String
Dim lsCuentaInt As String
Dim lsSubCuenta As String
Dim lsSubCuentaInt As String

Set oCtaIf = New NCajaCtaIF
Set oDMov = New DMov
On Error GoTo GrabaConfAperturaErr
nImporteLP = 0
lnFactor = 1

If Val(Left(psCtaIFCod, 2)) = gTpoCtaIFCtaAdeud Then
   nImporteLP = CalculaSaldoAdeudadoLargoPlazo(GetFechaMov(psMovNro, True), psPersCod, psIFTpo, psCtaIFCod)
   lnFactor = -1
End If

GrabaConfApertura2 = 1
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpecod, psMovDesc
lnMovNro = oDMov.GetnMovNro(psMovNro)
ldFecha = CDate(GetFechaMov(psMovNro, True))

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta de Bancos en el debe : Primer Capital
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe & psSubCtaDebe, (pnImporte - Format(nImporteLP * IIf(pnTasaVac > 0, pnTasaVac, 1), "#,#0.00")) * lnFactor
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod

'Creamos la SubCuenta de la Inst.Financ en CtaIFFiltro
oDMov.InsertaCuentaIFFiltro2 psPersCod, psIFTpo, psCtaIFCod, psCtaDebe, psSubCtaDebe
'Creamos la SubCuenta de la Inst.Financ en CtaIFFiltro - para intereses 5
If psPersCod = "1090100822183" And Mid(psCtaIFCod, 1, 2) = "04" Then 'PASI20140125 TI-ERS140-2013
    lsCuenta = "51M102"
Else
    lsCuenta = "51M103" & Format(psIFTpo, "00")
End If
'BCRP
If psPersCod = "1090100822183" Then
    lsCuentaInt = "11M802"
    If Left(psCtaIFCod, 2) = "03" Then 'DPF
        lsSubCuenta = "0103"
        lsSubCuentaInt = "01"
    Else 'Overnight
        'lsSubCuenta = "0104"
        lsSubCuenta = "09" 'PASI20140125 TI-ERS140-2013
        lsSubCuentaInt = "02"
    End If
Else
    lsSubCuenta = psSubCtaDebe
    lsCuentaInt = "11M803" & Format(psIFTpo, "00")
    lsSubCuentaInt = psSubCtaDebe
End If

oDMov.InsertaCuentaIFFiltro2 psPersCod, psIFTpo, psCtaIFCod, Replace(lsCuenta, "M", Mid(psCtaDebe, 3, 1)), lsSubCuenta
'Creamos la SubCuenta de la Inst.Financ en CtaIFFiltro - para intereses 11M8
oDMov.InsertaCuentaIFFiltro2 psPersCod, psIFTpo, psCtaIFCod, Replace(lsCuentaInt, "M", Mid(psCtaDebe, 3, 1)), lsSubCuentaInt

'ADEUDADOS Largo Plazo
If nImporteLP > 0 Then
    lnMovOrden = 0
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), "26" & Mid(psCtaDebe, 3, 20) & psSubCtaDebe, Format(nImporteLP * lnFactor * IIf(pnTasaVac > 0, pnTasaVac, 1), "#,#0.00")
    lnMovOrden = lnMovOrden + 1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod
    oDMov.ActualizaAdeudadosSaldoCap psPersCod, psIFTpo, psCtaIFCod, , nImporteLP
End If

'Guardamos la Cuenta de la Pendiente que se elimina
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, (pnImporte * -1) * lnFactor
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod

'Pasamos a activo el estado de la cuenta aperturada
oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIFCod, , , Format(ldFecha, gsFormatoFecha), Format(ldFecha, gsFormatoFecha), , , gEstadoCtaIFActiva
oDMov.InsertaMovRef lnMovNro, pnMovNroRef

If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"

oDMov.CommitTrans
lbTrans = False

GrabaConfApertura2 = 0
Set oDMov = Nothing

Exit Function
GrabaConfAperturaErr:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaConfApertura", lsMsgErr & " Error en Grabación de Confirmación de Apertura "
End Function
'END EJVG *******
'''''**

Public Function GrabaInteresDev(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
                                ByVal psCtaDebe As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                ByVal psPersCod As String, ByVal psIFTpo As String, _
                                ByVal psCtaIFCod As String, ByVal pdCtaIFInt As Date, ByVal pnDiasTrans As Integer) As Integer

Dim lnMovItem As Long
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
lnInteres = oCtaIf.GetInteresCtaIf(psPersCod, psIFTpo, psCtaIFCod)
lnInteres = lnInteres + pnImporte
Set oCtaIf = Nothing


oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpecod, psMovDesc
lnMovNro = oDMov.GetnMovNro(psMovNro)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1:
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, pnImporte
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod

lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnImporte * -1
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod

oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIFCod, , , , Format(pdCtaIFInt, gsFormatoFecha), , , , lnInteres, psMovNro
If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If
oDMov.CommitTrans
lbTrans = False

GrabaInteresDev = 0
Set oDMov = Nothing
Set oCtaIf = Nothing
Exit Function
GrabaInteresDevErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaInteresDev", Err.Description & " Error Grabación de Interes Devengado"
End Function
Public Function GrabaCapitalizaRegCheque(ByVal psMovNro As String, ByVal psOpecod As String, _
                                        ByVal psMovDesc As String, ByVal pdFechaMov As Date, _
                                        ByVal psCtaDebe As String, ByVal psObjCodProd As String, ByVal psAreaCod As String, ByVal psAgeCod As String, _
                                        ByVal psCtaHaberInteres As String, ByVal psCtaHaberIntCalc As String, _
                                        ByVal psCtaHaberCapital As String, ByVal pnImporteCapital As Currency, _
                                        ByVal pnImporteInt As Currency, ByVal pnImporteIntCalc As Currency, _
                                        ByVal rsObjHaberMot As ADODB.Recordset, _
                                        ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, ByVal psCtaIFCod As String, _
                                        ByVal psNumCheque As String, ByVal psPersCodIfChq As String, ByVal pnTipoIFChq As CGTipoIF, ByVal pnPlaza As ChequePlaza, _
                                        ByVal psCtaCheque As String, ByVal pnImporteTotal As Currency, ByVal pdFechaReg As Date, _
                                        ByVal pdFechaVal As Date, ByVal pnMonedaCheque As Moneda, _
                                        Optional ByVal pnEstadoCheque As ChequeEstado = gChqEstEnValorizacion, _
                                        Optional ByVal pnEstadoCG As CGEstadosChq = gCGEstadosChqRecibido, _
                                        Optional ByVal pnConfirmaCaja As CGEstadoConfCheque = ChqCGSinConfirmacion, _
                                        Optional ByVal psAreaCodChq As String, Optional ByVal psAgeCodChq As String, Optional pnDiasTrans As Integer = 0, Optional pnPlazo As Integer = 0, Optional pbCancela As Boolean = False) As Integer

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
oMov.BeginTrans
lbTrans = True
oMov.InsertaMov psMovNro, psOpecod, psMovDesc
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
'oMov.InsertaMovCtaIF lnMovNro, psPersCod, pnIFTpo, psCtaIfCod, CGCtaIFConInteres, (pnImporteIntCalc + pnImporteInt) * -1, pnDiasTrans

If pnImporteCapital > 0 And psCtaHaberCapital <> "" Then
    lnMovItem = lnMovItem + 1: lnMovOrden = 0
    lnMovOrden = lnMovOrden + 1
    oMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaberCapital, pnImporteCapital * -1
    oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, pnIFTpo, psCtaIFCod

    'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
    'oMov.InsertaMovCtaIF lnMovNro, psPersCod, pnIFTpo, psCtaIfCod, CGCtaIFConCapital, pnImporteCapital * -1
    'oMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                            psPersCod , pnIFTpo, psCtaIfCod
End If

'InsertaRefCap psPersCod, pnIFTpo, psCtaIfCod, lnMovNro, oMov
oMov.ActualizaCtaIF psPersCod, pnIFTpo, psCtaIFCod, , , Format(pdFechaMov, gsFormatoFecha), , pnPlazo, Format(pdFechaMov + pnPlazo, gsFormatoFecha), IIf(pbCancela, gEstadoCtaIFCancelada, -1), 0, psMovNro
'grabacion dentro de la tabla de cheques
oMov.InsertaCheque TpoDocCheque, psNumCheque, psPersCodIfChq, pnTipoIFChq, pnPlaza, psCtaCheque, pnImporteTotal, pdFechaReg, _
                    pdFechaVal, pnEstadoCG, pnConfirmaCaja, pnMonedaCheque, , psAreaCodChq, psAgeCodChq
oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIfChq, pnTipoIFChq, pdFechaReg, gChqEstEnValorizacion, psMovNro, psCtaCheque

If psNumCheque <> "" Then
    oMov.InsertaMovDoc lnMovNro, TpoDocCheque, psNumCheque, Format(pdFechaReg, gsFormatoFecha)
End If

If Mid(psOpecod, 3, 1) = "2" Then
    oMov.GeneraMovME lnMovNro, psMovNro
End If
oMov.ActualizaSaldoMovimiento psMovNro, "+"

oMov.CommitTrans
lbTrans = False
GrabaCapitalizaRegCheque = 0
Exit Function
GrabaCapitalizaRegChequeErr:
lsMsgErr = Err.Description
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise 50001, "nCajaGeneral: GrabaCapitalizaRegCheque", lsMsgErr
End Function
Private Sub InsertaRefCap(ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, _
                          ByVal psCtaIFCod As String, ByVal pnMovNro As Long, ByRef oMov As DMov)



sql = " SELECT  DISTINCT M.NMOVNRO, M.CMOVNRO , MC.NMOVIMPORTE" _
    & " FROM    MOV M " _
    & "         JOIN MOVCTAIF MC ON MC.NMOVNRO = M.NMOVNRO " _
    & "         JOIN CTAIF C     ON MC.CPERSCOD=C.CPERSCOD AND MC.CIFTPO = C.CIFTPO AND MC.cCtaIFCod=C.cCtaIFCod " _
    & " Where   MC.nConceptoCod =" & CGCtaIFConInteres & " And MC.NMOVIMPORTE > 0 And m.nMovFlag = " & gMovFlagVigente & " " _
    & "         AND SUBSTRING(M.CMOVNRO,1,8)> CONVERT(CHAR(8),dCtaIfCap,112) " _
    & "         AND MC.cPersCod='" & psPersCod & "' AND MC.cIFTpo='" & Format(pnIFTpo, "00") & "' AND MC.cCtaIFCod ='" & psCtaIFCod & "'"

Set rs = oMov.CargaRecordSet(sql)
Do While Not rs.EOF
    oMov.InsertaMovRef pnMovNro, rs!nMovNro
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
End Sub
Public Function GrabaCapitalizaIFEfectivo(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
                                         ByVal rsBillete As ADODB.Recordset, ByVal rsMoneda As ADODB.Recordset, _
                                         ByVal psCtaDebe As String, ByVal pnImporte As Currency, _
                                         ByVal psCtaDif As String, ByVal pnImporteDif As Currency, _
                                         ByVal psCtaHaberInt As String, ByVal pnImporteInt As Currency, _
                                         ByVal psCtaHaberIntCal As String, ByVal pnImporteIntCalc As Currency, _
                                         ByVal psCtaHaberCapital As String, ByVal pnImporteCapital As Currency, _
                                         ByVal psPersCod As String, ByVal psIFTpo As String, _
                                         ByVal psCtaIFCod As String, ByVal pdCtaIFCap As Date, Optional pnDiasTrans As Integer = 0, Optional pnPlazo As Integer = 0, Optional pbCancela As Boolean = False) As Integer
                    
Dim lnMovItem As Long
Dim lnMovOrden As Long
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

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpecod, psMovDesc
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
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDif, pnImporteDif
End If

If pnImporteInt > 0 Then
    lnMovItem = lnMovItem + 1:
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaberInt, pnImporteInt * -1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, psPersCod, psIFTpo, psCtaIFCod
End If
If pnImporteIntCalc > 0 Then
    lnMovItem = lnMovItem + 1:
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaberIntCal, pnImporteIntCalc * -1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, psPersCod, psIFTpo, psCtaIFCod
End If

'Referencia a todos los Movimientos de Intereses despues de la Ultima Capitalizacion
'InsertaRefCap psPersCod, psIFTpo, psCtaIfCod, lnMovNro, oDMov
If pnImporteCapital > 0 And psCtaHaberCapital <> "" Then
    lnMovItem = lnMovItem + 1:
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaberCapital, pnImporteCapital * -1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, psPersCod, psIFTpo, psCtaIFCod
    
    'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
    'oDMov.InsertaMovCtaIF lnMovNro, psPersCod, psIFTpo, psCtaIfCod, CGCtaIFConCapital, pnImporteCapital * -1
    'oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                            psPersCod , psIFTpo, psCtaIfCod
End If
oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIFCod, , , Format(pdCtaIFCap, gsFormatoFecha), , pnPlazo, Format(pdCtaIFCap + pnPlazo, gsFormatoFecha), IIf(pbCancela, gEstadoCtaIFCancelada, -1), 0, psMovNro

If Mid(psOpecod, 3, 1) = "2" Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"

oDMov.CommitTrans
lbTrans = False

GrabaCapitalizaIFEfectivo = 0
Set oDMov = Nothing
Set oContF = Nothing
Exit Function
GrabaCapitalizaIFEfectivoErr:
lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaCapitalizaIFEfectivo", lsMsgErr & " Error Graba Operaciones en Efectivo de Cuentas de IF"
End Function
Public Function GrabaCapitalizaDoc(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
                                    ByVal psCtaDebe As String, ByVal pnImporte As Currency, _
                                    ByVal pnTpoObj As TpoObjetos, ByVal psObjetoCodDebe As String, _
                                    ByVal psCtaHaberInt As String, ByVal pnImporteInt As Currency, _
                                    ByVal psCtaHaberIntCal As String, ByVal pnImporteIntCalc As Currency, _
                                    ByVal psCtaHaberCapital As String, ByVal pnImporteCapital As Currency, _
                                    ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psCtaIFCod As String, _
                                    ByVal pdCtaIFCap As Date, _
                                    ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pdFechaDoc As Date, ByVal psNroVoucher As String, _
                                    Optional ByVal psCtaAhorros As String, Optional ByVal pnMotivoNANC As MotivoNotaAbonoCargo = -1, Optional ByVal psObjetoPadre As String = "", Optional psObjeto As String = "", Optional pnDiasTrans As Integer = 0, Optional pnPlazo As Integer = 0, Optional pbCancela As Boolean = False) As Integer
            
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

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpecod, psMovDesc
lnMovNro = oDMov.GetnMovNro(psMovNro)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, pnImporte
lnMovOrden = lnMovOrden + 1
If pnTpoObj <> -1 Then
    Select Case pnTpoObj
        Case ObjEntidadesFinancieras
            oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, pnTpoObj
            oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psObjetoCodDebe, 4, 13), Mid(psObjetoCodDebe, 1, 2), Mid(psObjetoCodDebe, 18, 10)
            
            'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
            'oDMov.InsertaMovCtaIF lnMovNro, Mid(psObjetoCodDebe, 4, 13), Mid(psObjetoCodDebe, 1, 2), Mid(psObjetoCodDebe, 18, 10), CGCtaIFConCapital, pnImporte
            'oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
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

If pnImporteInt <> 0 Then
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaberInt, pnImporteInt * -1
End If
If pnImporteIntCalc <> 0 Then
    lnMovItem = lnMovItem + 1:
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaberIntCal, pnImporteIntCalc * -1
End If
''INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
'oDMov.InsertaMovCtaIF lnMovNro, psPersCod, psIFTpo, psCtaIfCod, CGCtaIFConInteres, (pnImporteInt + pnImporteIntCalc) * -1, pnDiasTrans

'InsertaRefCap psPersCod, psIFTpo, psCtaIfCod, lnMovNro, oDMov
oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIFCod, , , Format(pdCtaIFCap, gsFormatoFecha), , pnPlazo, Format(pdCtaIFCap + pnPlazo, gsFormatoFecha), IIf(pbCancela, gEstadoCtaIFCancelada, -1), 0, psMovNro
If pnImporteCapital > 0 And psCtaHaberCapital <> "" Then
    lnMovItem = lnMovItem + 1:
    lnMovOrden = 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaberCapital, pnImporteCapital * -1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod
    
    'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
    'oDMov.InsertaMovCtaIF lnMovNro, psPersCod, psIFTpo, psCtaIfCod, CGCtaIFConCapital, pnImporteCapital * -1
    'oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                            psPersCod , psIFTpo, psCtaIfCod
End If

If pnTpoDoc = TpoDocNotaAbono Or pnTpoDoc = TpoDocNotaCargo Then
    oDMov.ActualizaNotaAbonoCargo pnTpoDoc, psNroDoc, gNCNAEnMovimiento, , pnImporte
    oDMov.InsertaNotaAbonoCargoEst pnTpoDoc, psNroDoc, gNCNAEnMovimiento, psMovNro
End If
If psNroDoc <> "" Then
    oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, Format(pdFechaDoc, gsFormatoFecha)
End If
If psNroVoucher <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psNroVoucher, Format(pdFechaDoc, gsFormatoFecha)
End If

If Mid(psOpecod, 3, 1) = "2" Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If

oDMov.ActualizaSaldoMovimiento psMovNro, "+"
oDMov.CommitTrans
lbTrans = False

GrabaCapitalizaDoc = 0
Set oDMov = Nothing
Exit Function
GrabaCapitalizaDocErr:
lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaCapitalizaDoc", lsMsgErr & " Error Graba Capitalización con Documentos"
End Function


'NUEVAS FUNCIONES
'********************
Public Function GrabaPagoSunat(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
                              ByVal pnTpoDocPago As TpoDoc, ByVal psNroDocPago As String, ByVal psFechaDocPago As String, _
                              ByVal rsBillete As ADODB.Recordset, ByVal rsMoneda As ADODB.Recordset, _
                              ByVal rsDebe As ADODB.Recordset, _
                              ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                              ByVal psIFPersCod As String, ByVal psIFTpo As String, _
                              ByVal psCtaIFCod As String, _
                              ByVal pnTpoDocFormulario As TpoDoc, ByVal psNroDocFormulario As String, ByVal pdFechaDocFormulario As Date, _
                              ByVal psDocVoucher As String, Optional lsCtaITFD As String, Optional lsCtaITFH As String, Optional nTasaITF As Double = 0, Optional bAfectoITF As Boolean = True) As Integer
            
Dim lnMovItem  As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New DMov
On Error GoTo GrabaPagoSunatErr

GrabaPagoSunat = 1

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable
lnMovNro = oDMov.GetnMovNro(psMovNro)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"

'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
If Not rsDebe Is Nothing Then
   Do While Not rsDebe.EOF
      If rsDebe!Monto <> 0 Then
         lnMovItem = lnMovItem + 1
         oDMov.InsertaMovCta lnMovNro, lnMovItem, rsDebe!Cuenta, rsDebe!Monto
      End If
      rsDebe.MoveNext
   Loop
End If

Select Case pnTpoDocPago
   Case -1 'EFECTIVO
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
   Case TpoDocCarta, TpoDocCheque ''CUENTAS DE INSTITUCIONES FINANCIERAS
      lnMovItem = lnMovItem + 1: lnMovOrden = 1
      oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnImporte * -1
      'oDMov.InsertaMovCtaIF lnMovNro, psIFPersCod, psIFTpo, psCtaIfCod, CGCtaIFConCapital, pnImporte
      'oDMov.ActualizaSaldoCtaIF CDate(psFechaDocPago), psIFPersCod, psIFTpo, psCtaIfCod'EJVG20121108 Corregir Descuadres Saldos CtaIFSaldo
      oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
      oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psIFPersCod, psIFTpo, psCtaIFCod
      oDMov.InsertaMovDoc lnMovNro, pnTpoDocPago, psNroDocPago, psFechaDocPago
      oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocVoucher, psFechaDocPago
End Select

'Grabamos los datos del Formulario a Pagar
oDMov.InsertaMovDoc lnMovNro, pnTpoDocFormulario, psNroDocFormulario, Format(pdFechaDocFormulario, gsFormatoFecha)

'*---- John **********
If bAfectoITF = True Then 'bit si esta afecto al ITF el pago a Proveedor
    'Dim lnMontoITF As Currency
    Dim lnMontoITF As Double '*** PEAC 20110331
    If lsCtaITFD <> "" And lsCtaITFH <> "" And nTasaITF > 0 Then
        lnMontoITF = fgTruncar((pnImporte) * nTasaITF, 2)
        If lnMontoITF > 0 Then
            lnMovItem = lnMovItem + 1
            oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFD, lnMontoITF
            lnMovItem = lnMovItem + 1
            oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFH, lnMontoITF * -1
        End If
    End If
End If
'****************
oDMov.ActualizaSaldoMovimiento psMovNro, "+" 'EJVG20121108
oDMov.CommitTrans
lbTrans = False

GrabaPagoSunat = 0
Set oDMov = Nothing
Exit Function
GrabaPagoSunatErr:
   lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Set oDMov = Nothing
    Err.Raise vbObjectError + 100, "GrabaPagoSunat", lsMsgErr
End Function

Public Function GrabaCartaFianza(psMovNro As String, psOpecod As String, psMovDesc As String, pnTpoDocCartaFianza As TpoDoc, psNumCarta As String, psFechaEmision As String, _
                                psCodBanco As String, psCodPers As String, psFechaVenc As String, psCtaCodDebe As String, psCtaCodHaber As String, pnImporte As Currency) As Integer
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim oDMov As DMov
Set oDMov = New DMov
On Error GoTo GrabaCartaFianzaErr

GrabaCartaFianza = 1

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
oDMov.InsertaMovDoc lnMovNro, pnTpoDocCartaFianza, psNumCarta, psFechaEmision
oDMov.InsertaMovCta lnMovNro, 1, psCtaCodDebe, pnImporte
oDMov.InsertaMovCta lnMovNro, 2, psCtaCodHaber, pnImporte * -1
'oDMov.InsertaMovOtrosItem lnMovNro, 1, "CF", pnImporte, psCodBanco
'oDMov.InsertaMovOtrosItem lnMovNro, 2, "CF", pnImporte, Format(CDate(psFechaVenc), gsFormatoFecha)
oDMov.InsertaMovCartaFianza lnMovNro, psCodPers, psCodBanco, Format(CDate(psFechaVenc), gsFormatoFecha)
oDMov.InsertaMovGasto lnMovNro, psCodPers, "0"

If Mid(psOpecod, 3, 1) = "2" Then
   oDMov.GeneraMovME lnMovNro, psMovNro
End If
oDMov.CommitTrans
lbTrans = False

GrabaCartaFianza = 0
Set oDMov = Nothing
Exit Function
GrabaCartaFianzaErr:
   lsMsgErr = Err.Description
   If lbTrans Then
       oDMov.RollbackTrans
       lbTrans = False
   End If
   Set oDMov = Nothing
   Err.Raise vbObjectError + 100, "GrabaCartaFianza", lsMsgErr
End Function

Public Function GrabaCartaFianzaSalida(pdFecha As Date, psCodAge As String, psCodUser As String, psOpecod As String, psMovDesc As String, rsCarta As ADODB.Recordset, psCtaCodDebe As String, psCtaCodHaber As String) As String
Dim lbTrans  As Boolean
Dim lsMovNro As String
Dim lnMovNro As Long
Dim oDMov    As DMov
Dim aMov()   As String
Dim nMovs    As Integer
Dim k        As Integer

Set oDMov = New DMov
On Error GoTo GrabaCartaFianzaErr

GrabaCartaFianzaSalida = ""
If Not rsCarta Is Nothing Then
   If Not rsCarta.State = adStateClosed Then
      oDMov.BeginTrans
      lbTrans = True
      nMovs = 0
      Do While Not rsCarta.EOF
         If rsCarta!Itm = "1" Then
            lsMovNro = oDMov.GeneraMovNro(pdFecha, psCodAge, psCodUser)
            oDMov.InsertaMov lsMovNro, psOpecod, psMovDesc
            lnMovNro = oDMov.GetnMovNro(lsMovNro)
            oDMov.InsertaMovCont lnMovNro, rsCarta!Monto, 0, "0"
            oDMov.InsertaMovCta lnMovNro, 1, psCtaCodDebe, rsCarta!Monto
            oDMov.InsertaMovCta lnMovNro, 2, psCtaCodHaber, rsCarta!Monto * -1
            oDMov.InsertaMovRef lnMovNro, rsCarta!nMovNro
            If Mid(psOpecod, 3, 1) = "2" Then
               oDMov.GeneraMovME lnMovNro, lsMovNro
            End If
            nMovs = nMovs + 1
            ReDim Preserve aMov(4, nMovs)
            aMov(0, nMovs) = lsMovNro
            aMov(1, nMovs) = rsCarta!Persona
            aMov(2, nMovs) = rsCarta![Entidad Financiera]
            aMov(3, nMovs) = rsCarta![Nro Carta]
         End If
         rsCarta.MoveNext
      Loop
      oDMov.CommitTrans
      lbTrans = False
      Dim oFun   As NContImprimir
      Dim sImpre As String
      Dim lsImpreCabe As String
      Set oFun = New NContImprimir
      For k = 1 To nMovs
         lsImpreCabe = ""
         lsImpreCabe = lsImpreCabe & "  Proveedor           : " & aMov(1, k) & oImpresora.gPrnSaltoLinea
         lsImpreCabe = lsImpreCabe & "  Entiedad Financiera : " & aMov(2, k)
         sImpre = sImpre & oFun.ImprimeAsientoContable(aMov(0, k), 66, 80, "SALIDA DE CARTAS FIANZA Nro. " & aMov(3, k), lsImpreCabe, "19") & oImpresora.gPrnSaltoPagina
      Next
      GrabaCartaFianzaSalida = sImpre
      Set oFun = Nothing
   End If
End If
GrabaCartaFianzaSalida = sImpre
Set oDMov = Nothing
Exit Function
GrabaCartaFianzaErr:
   lsMsgErr = Err.Description
   If lbTrans Then
       oDMov.RollbackTrans
       lbTrans = False
   End If
   Set oDMov = Nothing
   Err.Raise vbObjectError + 100, "GrabaCartaFianza", lsMsgErr
End Function

'Public Function GetDatosCartaFianza(psCtaContCod As String, Optional pnTipo As Integer = 0, Optional psFechaDel As String = "", Optional psFechaAl As String = "") As ADODB.Recordset
'Dim lsFiltro As String
'
''Modificado por Pepe Funciona para Trujillo y Lima
'Select Case pnTipo
'   Case 0   'Todas las Cartas Vigentes
'      lsFiltro = " and Ref.nMovNro is NULL "
'      If psFechaDel <> "" Then
'         lsFiltro = lsFiltro & " and LEFT(M.cMovNro,8) BETWEEN '" & psFechaDel & "' and '" & psFechaAl & "' "
'      End If
'   Case 1   'Cartas Recibidas según Rango de Fechas
'      lsFiltro = " and LEFT(M.cMovNro,8) BETWEEN '" & psFechaDel & "' and '" & psFechaAl & "' "
'   Case 2   'Cartas Entregadas según Rango de Fechas
'      lsFiltro = " and LEFT(Ref.cMovNro,8) BETWEEN '" & psFechaDel & "' and '" & psFechaAl & "' "
'End Select
'oConect.AbreConexion
'sql = " SELECT MD.CDOCNRO, CONVERT(CHAR(30),M.CMOVDESC) DESCRIPCION , Bco.cPersNombre cBanco, P.cPersNombre, " _
'    & "        ISNULL(MME.NMOVMEIMPORTE, MC.NMOVIMPORTE) *-1  NMOVIMPORTE , MD.DDOCFECHA, CONVERT(datetime,MCF.dCartaVenc) AS FECHAVENC, MCF.cIFPersCod, substring(M.cmovnro,7,2) + '/' + substring(M.cmovnro,5,2) + '/' + substring(M.cmovnro,1,4) as dFecIng, M.cMovNro, M.nMovNro, ISNULL(Ref.cMovNro,'') cMovNroSalida " _
'    & " FROM MOVCTA MC INNER JOIN MOV M ON M.nMOVNRO=MC.nMOVNRO " _
'    & "    INNER JOIN MOVDOC       MD  ON MD.nMOVNRO=M.nMOVNRO " _
'    & "    INNER JOIN MOVCARTAFIANZA MCF ON MCF.nMOVNRO=M.nMOVNRO " _
'    & "    INNER JOIN Persona BCO ON BCO.CPERSCOD = MCF.cIFPersCod " _
'    & "    INNER JOIN Persona P   ON P.cPersCod=MCF.cPersCod " _
'    & "     LEFT JOIN (SELECT MR.nMOVNRO, MR.nMovNroRef, M1.cMovNro FROM MOVREF MR JOIN Mov M1 ON M1.nMovNro = mr.nMovNro " _
'    & "           WHERE  M1.nMOVESTADO= " & gMovEstContabMovContable & " AND M1.nMOVFLAG NOT  IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
'    & "          ) Ref ON Ref.nMovNroRef = m.nMovNro " _
'    & "     LEFT JOIN MOVME MME ON MME.nMOVNRO=MC.nMOVNRO AND MC.nMOVITEM=MME.nMOVITEM " _
'    & " WHERE MC.NMOVIMPORTE < 0 and MC.CCTACONTCOD = '" & Trim(psCtaContCod) & "' AND M.nMOVESTADO IN (" & gMovEstContabMovContable & "," & gMovEstContabNoContable & ") AND M.nMOVFLAG NOT  IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
'    & lsFiltro & " ORDER BY p.cPersNombre, m.cMovNro"
'    '
'Set GetDatosCartaFianza = oConect.CargaRecordSet(sql)
'End Function

Public Function GetDatosCartaFianza(psCtaContCod As String, Optional pnTipo As Integer = 0, Optional psFechaDel As String = "", Optional psFechaAl As String = "") As ADODB.Recordset
Dim lsFiltro As String

'Modificado por Pepe Funciona para Trujillo y Lima
Select Case pnTipo
   Case 0   'Todas las Cartas Vigentes
      lsFiltro = " and Ref.nMovNro is NULL "
      If psFechaDel <> "" Then
         lsFiltro = lsFiltro & " and LEFT(M.cMovNro,8) BETWEEN '" & psFechaDel & "' and '" & psFechaAl & "' "
      End If
   Case 1   'Cartas Recibidas según Rango de Fechas
      lsFiltro = " and LEFT(M.cMovNro,8) BETWEEN '" & psFechaDel & "' and '" & psFechaAl & "' "
   Case 2   'Cartas Entregadas según Rango de Fechas
      lsFiltro = " and LEFT(Ref.cMovNro,8) BETWEEN '" & psFechaDel & "' and '" & psFechaAl & "' "
End Select
oConect.AbreConexion
sql = " SELECT MD.CDOCNRO, CONVERT(CHAR(30),M.CMOVDESC) DESCRIPCION , Bco.cPersNombre cBanco, P.cPersNombre, " _
    & "        ISNULL(MME.NMOVMEIMPORTE, MC.NMOVIMPORTE) *-1  NMOVIMPORTE , MD.DDOCFECHA, CONVERT(datetime,MCF.dCartaVenc) AS FECHAVENC, MCF.cIFPersCod, substring(M.cmovnro,7,2) + '/' + substring(M.cmovnro,5,2) + '/' + substring(M.cmovnro,1,4) as dFecIng, M.cMovNro, M.nMovNro, ISNULL(Ref.cMovNro,'') cMovNroSalida " _
    & " FROM MOVCTA MC INNER JOIN MOV M ON M.nMOVNRO=MC.nMOVNRO " _
    & "    INNER JOIN MOVDOC       MD  ON MD.nMOVNRO=M.nMOVNRO " _
    & "    INNER JOIN MOVCARTAFIANZA MCF ON MCF.nMOVNRO=M.nMOVNRO " _
    & "    INNER JOIN Persona BCO ON BCO.CPERSCOD = MCF.cIFPersCod " _
    & "    INNER JOIN Persona P   ON P.cPersCod=MCF.cPersCod " _
    & "     LEFT JOIN (SELECT MR.nMOVNRO, MR.nMovNroRef, M1.cMovNro FROM MOVREF MR JOIN Mov M1 ON M1.nMovNro = mr.nMovNro " _
    & "           WHERE  M1.nMOVESTADO= " & gMovEstContabMovContable & " AND M1.nMOVFLAG NOT  IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
    & "          ) Ref ON Ref.nMovNroRef = m.nMovNro " _
    & "     LEFT JOIN MOVME MME ON MME.nMOVNRO=MC.nMOVNRO AND MC.nMOVITEM=MME.nMOVITEM " _
    & " WHERE MC.NMOVIMPORTE < 0 and MC.CCTACONTCOD = '" & Trim(psCtaContCod) & "' AND M.nMOVESTADO IN (" & gMovEstContabMovContable & "," & gMovEstContabNoContable & ") AND M.nMOVFLAG NOT  IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
    & lsFiltro & " ORDER BY p.cPersNombre, m.cMovNro"
    '
Set GetDatosCartaFianza = oConect.CargaRecordSet(sql)
End Function


Public Function GrabaCanjePorLetras(psMovNro As String, psOpecod As String, psMovDesc As String, rsLetras As ADODB.Recordset, psCtaContDebeB As String, psCtaContDebeS As String, pnImporteB As Currency, pnImporteS As Currency, rsDocs As ADODB.Recordset, psPersCod As String) As Boolean
Dim lbTrans   As Boolean
Dim lnMovNro  As Long
Dim lnMovItem As Integer
Dim oDMov     As DMov
Set oDMov = New DMov

Dim lnMovLetra As Long
Dim lsMovLetra As String

On Error GoTo GrabaCanjePorLetrasErr
GrabaCanjePorLetras = False

oDMov.BeginTrans
lbTrans = True

oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)

'Documentos Canjeados
If Not rsDocs Is Nothing Then
   Do While Not rsDocs.EOF
      If rsDocs!Itm = "1" Then
         oDMov.InsertaMovRef lnMovNro, rsDocs!nMovNro
         oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeB, 0, False
         oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeS, 0, False
      End If
      rsDocs.MoveNext
   Loop
End If

'Letras Generadas
If Not rsLetras.EOF Then
   Do While Not rsLetras.EOF
      lsMovLetra = oDMov.GeneraMovNro(, , , psMovNro)
      oDMov.InsertaMov lsMovLetra, psOpecod, "Letra " & rsLetras("Numero") & ". Canje de Documentos de Proveedor ", gMovEstContabPendiente, gMovFlagVigente
      lnMovLetra = oDMov.GetnMovNro(lsMovLetra)
      
      If pnImporteS <> 0 Then
         oDMov.InsertaMovCta lnMovLetra, 1, psCtaContDebeS, rsLetras("Monto") * -1
         oDMov.InsertaMovPendientesRend lnMovLetra, psCtaContDebeS, rsLetras("Monto")
      Else
         oDMov.InsertaMovCta lnMovLetra, 1, psCtaContDebeB, rsLetras("Monto") * -1
         oDMov.InsertaMovPendientesRend lnMovLetra, psCtaContDebeB, rsLetras("Monto")
      End If
      oDMov.InsertaMovDoc lnMovLetra, TpoDocLetras, rsLetras("Numero"), Format(rsLetras("Fec Vencim"), gsFormatoFecha)
      oDMov.InsertaMovOtrosItem lnMovLetra, 1, "L_Giro", 0, rsLetras("Fec Giro")
      oDMov.InsertaMovGasto lnMovLetra, psPersCod, ""
      
      rsLetras.MoveNext
   Loop
End If

oDMov.CommitTrans
lbTrans = False

GrabaCanjePorLetras = True
Exit Function
GrabaCanjePorLetrasErr:
   lsMsgErr = Err.Description
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    Set oDMov = Nothing
    Err.Raise vbObjectError + 100, "GrabaPagoProveedor", lsMsgErr
End Function

'ALPA 20090317*****************************************************
'Public Function GrabaPagoProveedor(psMovNro As String, psOpeCod As String, psMovDesc As String, psCtaContDebeB As String, psCtaContDebeS As String, _
'                                   psCtaContHaber As String, pnImporteB As Currency, pnImporteS As Currency, psPersCod As String, psTpoIf As String, psPersCodIF As String, psCtaIf As String, _
'                                   prsBilletaje As ADODB.Recordset, psDocTpo As String, psDocNro As String, psFechaDoc As String, psDocVoucher As String, rsDocs As ADODB.Recordset, psCuentaAho As String, _
'                                   ByVal psCtaDiferencia As String, ByVal pnMontoDiferencia As Currency, Optional ByVal gbBitCentral As Boolean, Optional pbGrabaOpeNegocio As Boolean = False, Optional pnRetAct As Currency = 0, Optional lsCtaITFD As String, Optional lsCtaITFH As String, Optional nTasaITF As Currency = 0, Optional bIncluyeNC As Boolean = False, _
'                                   Optional bAfectoITF As Boolean = True, Optional lnITFDescCuentas As Currency, Optional pnImporteRH As Currency, Optional psCtaContDebeRH As String, Optional psCtacontDebeRHJ As String, Optional pnImporteRHJ As Currency, Optional psCtaContDebeSegu As String, Optional pnImporteSegu As Currency, _
'                                   Optional psCtaContDebePagoVarios As String, Optional pnImportePagoVarios As Currency) As Integer
Public Function GrabaPagoProveedor(psMovNro As String, psOpecod As String, psMovDesc As String, psCtaContDebeB As String, psCtaContDebeS As String, _
                                   psCtaContHaber As String, pnImporteB As Currency, pnImporteS As Currency, psPersCod As String, psTpoIF As String, psPersCodIf As String, psCtaIf As String, _
                                   prsBilletaje As ADODB.Recordset, psDocTpo As String, psDocNro As String, psFechaDoc As String, psDocVoucher As String, rsDocs As ADODB.Recordset, psCuentaAho As String, _
                                   ByVal psCtaDiferencia As String, ByVal pnMontoDiferencia As Currency, Optional ByVal gbBitCentral As Boolean, Optional pbGrabaOpeNegocio As Boolean = False, Optional pnRetAct As Currency = 0, Optional lsCtaITFD As String, Optional lsCtaITFH As String, Optional nTasaITF As Double = 0, Optional bIncluyeNC As Boolean = False, _
                                   Optional bAfectoITF As Boolean = True, Optional lnITFDescCuentas As Double, Optional pnImporteRH As Currency, Optional psCtaContDebeRH As String, Optional psCtacontDebeRHJ As String, Optional pnImporteRHJ As Currency, Optional psCtaContDebeSegu As String, Optional pnImporteSegu As Currency, _
                                   Optional psCtaContDebePagoVarios As String, Optional pnImportePagoVarios As Currency, Optional nTipoPago As Integer = 0, Optional ByVal psCtaSaf As String = "", Optional ByVal lsCtaCodCore As String = "", Optional ByVal psCtaContLeasing As String = "") As Integer
                                   'Optional psCtaContDebePagoVarios As String, Optional pnImportePagoVarios As Currency, Optional nTipoPago As Integer = 0) As Integer
'******************************************************************
Dim lnMovItem  As Long
Dim lnMovOrden As Integer
Dim lnMovNro As Long
Dim lbTrans  As Boolean
Dim lsMovNro As String
Dim lsCadBol As String
Dim oDMov    As DMov
Set oDMov = New DMov

Dim oConst As NConstSistemas
Set oConst = New NConstSistemas
Dim oImp As DImpuesto
Set oImp = New DImpuesto
Dim oFun As NContFunciones
Set oFun = New NContFunciones
Dim lbBitReten As Boolean
Dim lsCtaReten As String
Dim lsCtaRetenConvesion As String
Dim lsCtaRetenCambio As String
Dim lbBCAR As Boolean
Dim lnTasaImp As Currency
Dim lnIngresos As Currency
Dim lnRetencion As Currency
Dim lnTopeRetencion As Currency
Dim lnRetAct As Currency
Dim lnRetActME As Currency
Dim lsRetDocNro As String
Dim lMN As Boolean
Dim lsMovNroNegocio As String
Dim lbGraboNegocio As Boolean

On Error GoTo GrabaPagoProveedorErr

GetTipCambio GetFechaMov(psMovNro, True)

If Mid(psCtaContHaber, 3, 1) = Moneda.gMonedaExtranjera Then
    lMN = False
Else
    lMN = True
End If

lbBitReten = IIf(oConst.LeeConstSistema(gConstSistBitRetencion6Porcent) = 1, True, False)
If lbBitReten And pnRetAct <> 0 Then
   lsRetDocNro = oFun.GeneraDocNro(TpoDocAlmacenComprdeRetencion6)
    lsCtaReten = oConst.LeeConstSistema(gConstSistCtaRetencion6Porcent)
End If

oDMov.BeginTrans
lbTrans = True
GrabaPagoProveedor = 1

oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu, "0", "0"

'ALPA 20110909*****************************************************
If Len(Trim(psCtaSaf)) > 0 Then
Dim oCtSaldo As DCtaSaldo
Set oCtSaldo = New DCtaSaldo
Call oCtSaldo.InsertarDesembolso_Log(psCtaSaf, lsCtaCodCore, lnMovNro)
End If
'******************************************************************

'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
If pnImporteB <> 0 Then
   lnMovItem = lnMovItem + 1
   'ALPA 20090317******************************************************
   'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeB, pnImporteB
   'ALPA 20110912
   'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeB, pnImporteB, nTipoPago
   If Len(Trim(psCtaSaf)) = 0 Then
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeB, pnImporteB, nTipoPago
   Else
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContLeasing, pnImporteB, nTipoPago
   End If
   '*******************************************************************
End If
If pnImporteS <> 0 Then
   lnMovItem = lnMovItem + 1
   'ALPA 20090317******************************************************
   'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeS, pnImporteS
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeS, pnImporteS, nTipoPago
   '*******************************************************************
End If
'JEOM
If pnImporteRH <> 0 Then
   lnMovItem = lnMovItem + 1
   'ALPA 20090317******************************************************
   'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeRH, pnImporteRH
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeRH, pnImporteRH, nTipoPago
   '*******************************************************************
End If
If pnImporteRHJ <> 0 Then
   lnMovItem = lnMovItem + 1
   'ALPA 20090317******************************************************
   'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtacontDebeRHJ, pnImporteRHJ
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtacontDebeRHJ, pnImporteRHJ, nTipoPago
   '*******************************************************************
End If
If pnImporteSegu <> 0 Then
   lnMovItem = lnMovItem + 1
   'ALPA 20090317******************************************************
   'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeSegu, pnImporteSegu
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeSegu, pnImporteSegu, nTipoPago
   '*******************************************************************
End If

If pnImportePagoVarios <> 0 Then
   lnMovItem = lnMovItem + 1
   'ALPA 20090317*******************************************************
   'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebePagoVarios, pnImportePagoVarios
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebePagoVarios, pnImportePagoVarios, nTipoPago
   '********************************************************************
End If

'Agregado

lnRetAct = 0
lnRetActME = 0
If lbBitReten Then
   lnRetAct = pnRetAct
   ''   lnRetActME = Round(lnRetAct / gnTipCambioPonderado, 2)
   lnRetActME = Round(lnRetAct / gnTipCambioPonderado, 2)
   lnMovItem = lnMovItem + 1
   'ALPA 20090317*********************************************************
   'oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCtaReten, lnRetAct * -1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCtaReten, lnRetAct * -1, nTipoPago
   oDMov.InsertaMovDoc lnMovNro, TpoDocAlmacenComprdeRetencion6, lsRetDocNro, psFechaDoc
   '***********************************************************************
End If

'guardamos el billetaje si este posee
If Not prsBilletaje Is Nothing Then
   oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, prsBilletaje, psCtaContHaber, "H"
Else
    'guardamos la cuenta del haber
'    If (pnImporteB + pnImporteS - lnRetAct) <> 0 Then
    If (pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu + pnImportePagoVarios - IIf(lMN, lnRetAct, lnRetActME)) <> 0 Then
        lnMovItem = lnMovItem + 1: lnMovOrden = 0
'        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber, (pnImporteB + pnImporteS - IIf(lnRetAct > 0, lnRetAct, 0)) * -1
        'ALPA 20090317*******************************************************
        'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber, (pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu + pnImportePagoVarios - IIf(lnRetAct + lnRetActME > 0, IIf(lMN, lnRetAct, lnRetActME), 0)) * -1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber, (pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu + pnImportePagoVarios - IIf(lnRetAct + lnRetActME > 0, IIf(lMN, lnRetAct, lnRetActME), 0)) * -1, nTipoPago
        '********************************************************************
    End If
    If psPersCodIf <> "" Then
        'guardamos la Cta de la IF si es que utiliza esta
        lnMovOrden = lnMovOrden + 1
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCodIf, psTpoIF, psCtaIf
        
    End If
End If
If psDocNro <> "" Then
   oDMov.InsertaMovDoc lnMovNro, psDocTpo, psDocNro, psFechaDoc
End If
If psDocVoucher <> "" Then
   oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocVoucher, GetFechaMov(psMovNro, False)
End If

oDMov.InsertaMovGasto lnMovNro, psPersCod, "0"

If pnMontoDiferencia <> 0 Then
    lnMovItem = lnMovItem + 1
    'oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDiferencia, pnMontoDiferencia * -1
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDiferencia, pnMontoDiferencia * -1, nTipoPago
End If

'Referenciamos todos los Documentos que se Pagan
If Not rsDocs Is Nothing Then '*** PEAC 20110218
If bIncluyeNC Or rsDocs!nDocTpo = 7 Then
    '   OJO: Pago Total
'    If Not rsDocs Is Nothing Then
       Do While Not rsDocs.EOF
          If rsDocs!Itm = "1" Then
            'ALPA 20090317***********************************************
             'oDMov.InsertaMovRef lnMovNro, rsDocs!nMovNro
             oDMov.InsertaMovRef lnMovNro, rsDocs!nMovNro, , nTipoPago
             '***********************************************************
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeB, 0, False
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeS, 0, False
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeRH, 0, False
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtacontDebeRHJ, 0, False
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeSegu, 0, False
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebePagoVarios, 0, False
         
          End If
          rsDocs.MoveNext
       Loop
    'End If
Else
    ' Puede ser pago parcial
    
    Dim nMontoPago As Currency
    'If Not rsDocs Is Nothing Then
        nMontoPago = pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu + pnImportePagoVarios
       
       'Do While Not rsDocs.EOF And nMontoPago > 0 '*** PEAC 20110301
       Do While Not rsDocs.EOF
       
          If rsDocs!Itm = "1" Then
             'ALPA 20090317************************************************
             'oDMov.InsertaMovRef lnMovNro, rsDocs!nMovNro
             oDMov.InsertaMovRef lnMovNro, rsDocs!nMovNro, , nTipoPago
             '*************************************************************
             If nMontoPago >= rsDocs!Importe Then
             
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeB, 0, False
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeS, 0, False
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeRH, 0, False
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtacontDebeRHJ, 0, False
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeSegu, 0, False
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebePagoVarios, 0, False
                nMontoPago = nMontoPago - rsDocs!Importe
             Else
                If pnImporteB <> 0 Then
                    oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebeB, rsDocs!Importe - nMontoPago
                Else
                   If pnImporteS <> 0 Then
                      'oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebeS, rsDocs!Importe - nMontoPago '*** PEAC 20110301
                      oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebeS, rsDocs!Importe - rsDocs!VALORSOLES
                   Else
                      If pnImporteRH <> 0 Then
                        oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebeRH, rsDocs!Importe - nMontoPago
                      Else
                        If pnImporteRHJ <> 0 Then
                           oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtacontDebeRHJ, rsDocs!Importe - nMontoPago
                        Else
                            If pnImportePagoVarios <> 0 Then
                               oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebePagoVarios, rsDocs!Importe - nMontoPago
                            Else
                               oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebeSegu, rsDocs!Importe - nMontoPago
                            End If
                        End If
                      End If
                   End If
                End If
                'nMontoPago = 0 '*** PEAC 20110301
             End If
          End If
          rsDocs.MoveNext
       Loop
    'End If

End If
End If

'Fin Agregado

'--Calculo del ITF
'--  JOHN
If bAfectoITF = True Then 'bit si esta afecto al ITF el pago a Proveedor
    'Dim lnMontoITF As Currency
    Dim lnMontoITF As Double '*** PEAC 20110331
    If lsCtaITFD <> "" And lsCtaITFH <> "" And nTasaITF > 0 Then
        lnMontoITF = fgTruncar((pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu + pnImportePagoVarios - lnRetAct) * nTasaITF, 2)
        If lnMontoITF > 0 Then
            lnMovItem = lnMovItem + 1
            'ALPA 20090317*********************************************
            'oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFD, lnMontoITF
            oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFD, lnMontoITF, nTipoPago
            '**********************************************************
            lnMovItem = lnMovItem + 1
            'ALPA 20090317*********************************************
            'oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFH, lnMontoITF * -1
            oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFH, lnMontoITF * -1, nTipoPago
            '**********************************************************
        End If
    End If
End If
'----------

If psOpecod <> "402581" And psOpecod <> "402582" Then
    If Mid(psOpecod, 3, 1) = 2 Then
        If lbBitReten Then
            'ALPA 20090430****************************
            'oDMov.GeneraMovME lnMovNro, psMovNro, gnTipCambioPonderado
            oDMov.GeneraMovME lnMovNro, psMovNro, gnTipCambioPonderado, , , nTipoPago
            '*****************************************
        Else
            'ALPA 20090430****************************
            'oDMov.GeneraMovME lnMovNro, psMovNro
            oDMov.GeneraMovME lnMovNro, psMovNro, , , , nTipoPago
            '********************************
        End If
    End If
End If

oDMov.ActualizaSaldoMovimiento psMovNro, "+"

If psDocTpo = TpoDocNotaAbono Or psDocTpo = TpoDocNotaCargo Then
    oDMov.InsertaRegDocCuenta psDocTpo, psDocNro, psCuentaAho
    lbGraboNegocio = False
    If pbGrabaOpeNegocio Then
        Dim oCapta As NCapMovimientos
        Dim lnSaldo As Double
        Set oCapta = New NCapMovimientos
        Dim oCon As New DConecta
        Dim oDis As New NRHProcesosCierre
        If gbBitCentral Then
           lsMovNroNegocio = oDMov.GeneraMovNro(, , , psMovNro)
           lbGraboNegocio = True
           lnSaldo = oCapta.CapAbonoCuentaAho(psCuentaAho, pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu + pnImportePagoVarios - lnRetAct, gCGPagProvAbonoCent, lsMovNroNegocio, psMovDesc, TpoDocNotaAbono, psDocNro, , , , , , , , , , True, False, , , oDMov.GetConexion, False, bAfectoITF, CDbl(lnMontoITF))
        Else
            If oCon.AbreConexion(Left(psCuentaAho, 2)) Then
                oCon.BeginTrans
                lbGraboNegocio = True
                lnSaldo = oDis.Abono(psCuentaAho, pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu + pnImportePagoVarios - lnRetAct, gCGPagProvAbonoDist, gCGPagProvAbonoDist, "112" & Left(psCuentaAho, 2), Right(psMovNro, 4), psDocNro, "PAGO : " & psMovDesc, oCon, GetFechaMov(psMovNro, True))
                oCon.CommitTrans
            End If
            oCon.CierraConexion
            Set oCon = Nothing
            Set oDis = Nothing
        End If
    End If
End If

oDMov.CommitTrans
lbTrans = False
GrabaPagoProveedor = 0


Set oDMov = Nothing

Exit Function
GrabaPagoProveedorErr:
   lsMsgErr = Err.Description
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    If lbGraboNegocio Then
        lsMsgErr = lsMsgErr & Chr(10) & Chr(10) & "EL SISTEMA HIZO EL ABONO EN LA CUENTA DEL CLIENTE. POR FAVOR EXTORNAR ESTA OPERACION"
    End If
    Set oDMov = Nothing
    Err.Raise vbObjectError + 100, "GrabaPagoProveedor", lsMsgErr
End Function

Public Function GrabaPagoProveedorEntregaDoc(psMovNro As String, psOpecod As String, psMovDesc As String, prs As ADODB.Recordset) As Integer
Dim lnMovItem  As Integer
Dim lnMovNro As Long
Dim oDMov    As DMov
Dim lbTrans  As Boolean
Dim lsMovNro As String
Set oDMov = New DMov
On Error GoTo GrabaPagoProveedorErr
oDMov.BeginTrans
lbTrans = True
GrabaPagoProveedorEntregaDoc = 1

oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
If Not prs Is Nothing Then
   Do While Not prs.EOF
      If prs!Itm = "1" Then
         oDMov.InsertaMovRef lnMovNro, prs!nMovNro
      End If
      prs.MoveNext
   Loop
End If
oDMov.CommitTrans
lbTrans = False
GrabaPagoProveedorEntregaDoc = 0

Set oDMov = Nothing

Exit Function
GrabaPagoProveedorErr:
   lsMsgErr = Err.Description
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    Set oDMov = Nothing
    Err.Raise vbObjectError + 100, "GrabaPagoProveedorEntregaDoc", lsMsgErr
End Function

Public Function GetDatosPagoProveedor(psCtaContCod As String, pnDocTpoPago As TpoDoc, Optional pnTipo As Integer = 0, Optional psFechaDel As String = "", Optional psFechaAl As String, Optional pbHaber As Boolean = True) As ADODB.Recordset
Dim sql      As String
Dim lsFiltro As String

On Error GoTo GetDatosPagoProveedorErr
Select Case pnTipo
   Case 0   'Todas las OP emitidas no Entregadas
      lsFiltro = " and     Entr.nMovNro is NULL "
   Case 1   'Todas las OP Entregadas no Pagadas
      lsFiltro = " and Not Entr.nMovNro is NULL and     Pago.nMovNro is NULL "
   Case 2   'Todas las OP Pagadas
      lsFiltro = " and Not Entr.nMovNro is NULL and Not Pago.nMovNro is NULL "
End Select

'FALTA DEFINIR OPERACIONES DE PAGO
sql = " SELECT  md.dDocFecha, d.cDocAbrev, md.nDocTpo, md.cDocNro, p.cPersNombre, m.cMovDesc, " _
    & "         mg.cPersCod,  m.cMovNro, m.nMovNro, IsNull(me.nMovMEImporte,mc.nMovImporte) * -1 as nDocImporte " _
    & " FROM    Mov m " _
    & "         JOIN MovCta   mc ON mc.nMovNro = m.nMovNro " _
    & "         JOIN MovDoc   md ON md.nMovNro = m.nMovNro JOIN Documento d ON d.nDocTpo = md.nDocTpo " _
    & "    LEFT JOIN MovMe    me ON me.nMovNro =mc.nMovNro and me.nMovItem=mc.nMovItem " _
    & "         JOIN MovGasto mg ON mg.nMovnro =mc.nMovNro " _
    & "         JOIN Persona  p  ON p.cPersCod = mg.cPersCod " _
    & "    LEFT JOIN (SELECT mr.nMovNro, mr.nMovNroRef, m.cMovNro FROM MovRef mr JOIN Mov m ON m.nMovNro = mr.nMovNro " _
    & "               WHERE m.nMovFlag = " & gMovFlagVigente & " and m.cOpeCod IN ('" & OpeCGOpeProvEntrOP & "','" & OpeCGOpeProvEntrCh & "','" & OpeCGOpeProvEntrOPME & "','" & OpeCGOpeProvEntrChME & "') " _
    & "              ) Entr ON Entr.nMovNroRef = m.nMovNro " _
    & "    LEFT JOIN (SELECT mr.nMovNro, mr.nMovNroRef, m.cMovNro FROM MovRef mr JOIN Mov m ON m.nMovNro = mr.nMovNro " _
    & "               WHERE m.nMovFlag = " & gMovFlagVigente & " and m.cOpeCod IN ('') " _
    & "              ) Pago ON Pago.nMovNroRef = m.nMovNro " _
    & " WHERE   mc.cCtaContCod IN ('" & psCtaContCod & "') and LEFT(m.cMovNro,8) BETWEEN '" & psFechaDel & "' and '" & psFechaAl & "' " _
    & "         and  md.nDocTpo = '" & pnDocTpoPago & "' and mc.nMovImporte " & IIf(pbHaber, "<", ">") & " 0 " & lsFiltro _
    & " ORDER BY md.cDocNro, p.cPersNombre "
    
Set GetDatosPagoProveedor = oConect.CargaRecordSet(sql)

Exit Function
GetDatosPagoProveedorErr:
    Err.Raise vbObjectError + 100, "GetDatosPagoProveedor", Err.Description
End Function


Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oIni = New ClasIni
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing
lbConexionPropia = True
Set oConect = New DConecta
oConect.AbreConexion
End Sub

Private Sub Class_Terminate()
If lbConexionPropia = True Then
    oConect.CierraConexion
    Set oConect = Nothing
End If
End Sub

'--- NUEVAS FUNCIONES
Public Function GetCajaRegOpeVentanilla(psOpecod As String, pdFechaDel As Date, pdFechaAl As Date) As ADODB.Recordset
Dim sql As String
On Error GoTo GetCajaRegOpeVentanillaErr
sql = "SELECT Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103) dFecha, p.cPersNombre,ISNULL(mr.nMovNroRef,'') nNumTran, " & IIf(Mid(psOpecod, 3, 1) = "2", "me.nMovMEImporte", "mc.nMovImporte") & "* -1, m.cMovDesc, mg.cPersCod, m.cMovNro, m.nMovNro, ISNULL(mr.cAgeCodRef,'') cAgeCodRef " _
     & "FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro " _
     & IIf(Mid(psOpecod, 3, 1) = "2", " JOIN MovMe  me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
     & "           JOIN MovGasto mg ON mg.nMovNro = m.nMovNro " _
     & "          JOIN Persona p ON p.cPersCod = mg.cPersCod LEFT JOIN MovRef mr ON mr.nMovNro = m.nMovNro " _
     & "WHERE m.cOpeCod = '" & psOpecod & "' and m.nMovEstado = " & gMovEstContabMovContable & " and m.nMovFlag = " & gMovFlagVigente _
     & "  and LEFT(m.cMovNro,8) BETWEEN '" & Format(pdFechaDel, gsFormatoMovFecha) & "' and '" & Format(pdFechaAl, gsFormatoMovFecha) & "' " _
     & "  and mc.nMovImporte < 0 "
Set GetCajaRegOpeVentanilla = oConect.CargaRecordSet(sql)
Exit Function
GetCajaRegOpeVentanillaErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetCajaBancosOperaciones(psOpecod As String, pdFechaDel As Date, pdFechaAl As Date, Optional pbVerReferencia As Boolean = False, Optional psOpeCodRef As String = "", Optional pbVerRefNeg As Boolean = False, Optional psVerMonto As String = "", Optional ByVal psOpeCodExt As String) As ADODB.Recordset
On Error GoTo GetCajaBancosOperacionesErr
sql = "SELECT Distinct Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103) dFecha, p.cPersNombre, cif.cCtaIFDesc, ISNULL(d.cDocAbrev,''), ISNULL(md.cDocNro,'Ninguno'), Abs(" & IIf(Mid(psOpecod, 3, 1) = "2", "ISNULL(me.nMovMEImporte,0)", "mc.nMovImporte") & "), m.cMovDesc, cif.cPersCod, cif.cIFTpo, cif.cCtaIFCod, m.cMovNro, m.nMovNro, ISNULL(md.nDocTpo,0) " & IIf(pbVerRefNeg, ", ISNULL(mrN.nMovNroRef,0) Ope_Age, ISNULL(mrN.cAgeCodRef,'') Agencia ", "") & ", nImporteIntDev = ISNULL( CASE WHEN m.cOpeCod IN ('" & gOpeCGOpeCancCtaPFMN & "','" & gOpeCGOpeCancCtaPFME & "') THEN (SELECT ISNULL(meI.nMovMEImporte,mcI.nMovImporte) * -1 FROM MovCta mcI LEFT JOIN MovMe meI ON meI.nMovNro = mcI.nMovNro and meI.nMovItem = mcI.nMovItem WHERE mcI.nMovNro = m.nMovNro and mcI.cCtaContCod like '11_8%') END, 0) " _
     & "FROM Mov m LEFT JOIN MovDoc md ON md.nMovNro = m.nMovNro and md.nDocTpo <>'56' LEFT JOIN Documento d ON d.nDocTpo = md.nDocTpo JOIN MovCta mc ON mc.nMovNro = m.nMovNro " _
     & IIf(Mid(psOpecod, 3, 1) = "2", " JOIN MovMe  me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
     & "           JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem " _
     & "           JOIN CtaIF cif ON cif.cIFTpo = mif.cIFTpo and cif.cPersCod = mif.cPersCod and cif.cCtaIFCod = mif.cCtaIFCod " _
     & "           JOIN Persona p ON p.cPersCod = mif.cPersCod " _
     & IIf(pbVerRefNeg, "LEFT JOIN MovRef mrN ON mrN.nMovNro = m.nMovNro and not RTRIM(ISNULL(mrN.cAgeCodRef,'')) = '' ", "") _
     & IIf(pbVerReferencia, " LEFT JOIN ( SELECT m1.nMovNro, mr.nMovNroRef FROM Mov m1 JOIN MovRef mr ON mr.nMovNro = m1.nMovNro WHERE m1.cOpecod = '" & psOpeCodRef & "' and m1.nMovEstado = " & gMovEstContabMovContable & " and m1.nMovFlag = " & gMovFlagVigente & " ) Ref ON Ref.nMovNroRef = m.nMovNro ", "") _
     & "WHERE m.cOpeCod IN ('" & psOpecod & "') and m.nMovEstado = " & gMovEstContabMovContable & " and m.nMovFlag = " & gMovFlagVigente _
     & "  and LEFT(m.cMovNro,8) BETWEEN '" & Format(pdFechaDel, gsFormatoMovFecha) & "' and '" & Format(pdFechaAl, gsFormatoMovFecha) & "' " _
     & IIf(pbVerReferencia, " and Ref.nMovNro IS NULL ", "") _
     & IIf(psVerMonto <> "", " and mc.nMovImporte " & psVerMonto & " 0", "") _
     & IIf(psOpeCodExt = "401561" Or psOpeCodExt = "402561" Or psOpeCodExt = "401661" Or psOpeCodExt = "402661", " And mif.nMovItem = 1", "")
Set GetCajaBancosOperaciones = oConect.CargaRecordSet(sql)
Exit Function
GetCajaBancosOperacionesErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'Public Function GetCajaOperacionesGen(psOpeCod As String, pdFechaDel As Date, pdFechaAl As Date, Optional pbVerReferencia As Boolean = False, Optional psOpeCodRef As String = "", Optional lnNroDigOperacion As Integer = 0) As ADODB.Recordset
'Dim lsOpeFiltro As String
'On Error GoTo GetCajaOperacionesGenErr
''lnNroDigOperacion = Nro de Caracteres del Codigo de Operación que se desea Buscar
'If lnNroDigOperacion > 0 Then
'    lsOpeFiltro = "LEFT(m.cOpeCod," & lnNroDigOperacion & ") = '" & Left(psOpeCod, lnNroDigOperacion) & "' "
'Else
'    lsOpeFiltro = "m.cOpeCod = '" & psOpeCod & "' "
'End If
'sql = "SELECT Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103) dFecha, ISNULL(d.cDocAbrev,''), ISNULL(md.cDocNro,'Ninguno'), Abs(" & IIf(Mid(psOpeCod, 3, 1) = "2", "SUM( ISNULL(me.nMovMEImporte,0) )", "SUM(mc.nMovImporte)") & "), mg.cPersCod, p.cPersNombre, m.cMovDesc, m.cMovNro, m.nMovNro, ISNULL(md.nDocTpo,0) " _
'    & "FROM Mov m LEFT JOIN MovDoc md ON md.nMovNro = m.nMovNro and not md.nDocTpo in (" & TpoDocVoucherEgreso & "," & TpoDocFormSUNAT & ") LEFT JOIN Documento d ON d.nDocTpo = md.nDocTpo JOIN MovCta mc ON mc.nMovNro = m.nMovNro " _
'    & IIf(Mid(psOpeCod, 3, 1) = "2", " JOIN MovMe  me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
'    & "      LEFT JOIN MovGasto mg ON mg.nMovNro = mc.nMovNro " _
'    & "      LEFT JOIN Persona p ON p.cPersCod = mg.cPersCod " _
'    & IIf(pbVerReferencia, " LEFT JOIN ( SELECT m1.nMovNro, mr.nMovNroRef FROM Mov m1 JOIN MovRef mr ON mr.nMovNro = m1.nMovNro WHERE m1.cOpecod = '" & psOpeCodRef & "' and m1.nMovEstado = " & gMovEstContabMovContable & " and m1.nMovFlag = " & gMovFlagVigente & " ) Ref ON Ref.nMovNroRef = m.nMovNro ", "") _
'    & "WHERE " & lsOpeFiltro & "and m.nMovEstado = " & gMovEstContabMovContable & " and m.nMovFlag = " & gMovFlagVigente & " and mc.nmovimporte > 0 " _
'    & "  and LEFT(m.cMovNro,8) BETWEEN '" & Format(pdFechaDel, gsFormatoMovFecha) & "' and '" & Format(pdFechaAl, gsFormatoMovFecha) & "' " _
'    & IIf(pbVerReferencia, " and Ref.nMovNro IS NULL ", "") _
'    & "GROUP BY Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103), " _
'    & "      ISNULL(d.cDocAbrev,''), ISNULL(md.cDocNro,'Ninguno'), mg.cPersCod, p.cPersNombre, m.cMovDesc, m.cMovNro, m.nMovNro, ISNULL(md.nDocTpo,0) " _
'    & " ORDER BY m.cMovNro "
'Set GetCajaOperacionesGen = oConect.CargaRecordSet(sql)
'Exit Function
'GetCajaOperacionesGenErr:
'    Err.Raise Err.Number, Err.Source, Err.Description
'End Function

Public Function GetCajaOperacionesGen(psOpecod As String, pdFechaDel As Date, pdFechaAl As Date, Optional pbVerReferencia As Boolean = False, Optional psOpeCodRef As String = "", Optional lnNroDigOperacion As Integer = 0) As ADODB.Recordset
Dim lsOpeFiltro As String
On Error GoTo GetCajaOperacionesGenErr
'lnNroDigOperacion = Nro de Caracteres del Codigo de Operación que se desea Buscar
If lnNroDigOperacion > 0 Then
    lsOpeFiltro = "LEFT(m.cOpeCod," & lnNroDigOperacion & ") = '" & Left(psOpecod, lnNroDigOperacion) & "' "
Else
    lsOpeFiltro = "m.cOpeCod = '" & psOpecod & "' "
End If
sql = "SELECT Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103) dFecha, ISNULL(d.cDocAbrev,''), ISNULL(md.cDocNro,'Ninguno'), Abs(" & IIf(Mid(psOpecod, 3, 1) = "2" And psOpecod <> 402581 And psOpecod <> 402582, "SUM( ISNULL(me.nMovMEImporte,0) )", "SUM(mc.nMovImporte)") & "), mg.cPersCod, p.cPersNombre, m.cMovDesc, m.cMovNro, m.nMovNro, ISNULL(md.nDocTpo,0) " _
    & "FROM Mov m LEFT JOIN MovDoc md ON md.nMovNro = m.nMovNro and not md.nDocTpo in (" & TpoDocVoucherEgreso & "," & TpoDocFormSUNAT & ") LEFT JOIN Documento d ON d.nDocTpo = md.nDocTpo JOIN MovCta mc ON mc.nMovNro = m.nMovNro " _
    & IIf(Mid(psOpecod, 3, 1) = "2" And psOpecod <> 402581 And psOpecod <> 402582, " JOIN MovMe  me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
    & "      LEFT JOIN MovGasto mg ON mg.nMovNro = mc.nMovNro " _
    & "      LEFT JOIN Persona p ON p.cPersCod = mg.cPersCod " _
    & IIf(pbVerReferencia, " LEFT JOIN ( SELECT m1.nMovNro, mr.nMovNroRef FROM Mov m1 JOIN MovRef mr ON mr.nMovNro = m1.nMovNro WHERE m1.cOpecod = '" & psOpeCodRef & "' and m1.nMovEstado = " & gMovEstContabMovContable & " and m1.nMovFlag = " & gMovFlagVigente & " ) Ref ON Ref.nMovNroRef = m.nMovNro ", "") _
    & "WHERE " & lsOpeFiltro & "and m.nMovEstado = " & gMovEstContabMovContable & " and m.nMovFlag = " & gMovFlagVigente & " and mc.nmovimporte > 0 " _
    & "  and LEFT(m.cMovNro,8) BETWEEN '" & Format(pdFechaDel, gsFormatoMovFecha) & "' and '" & Format(pdFechaAl, gsFormatoMovFecha) & "' " _
    & IIf(pbVerReferencia, " and Ref.nMovNro IS NULL ", "") _
    & "GROUP BY Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103), " _
    & "      ISNULL(d.cDocAbrev,''), ISNULL(md.cDocNro,'Ninguno'), mg.cPersCod, p.cPersNombre, m.cMovDesc, m.cMovNro, m.nMovNro, ISNULL(md.nDocTpo,0) " _
    & " ORDER BY m.cMovNro "
Set GetCajaOperacionesGen = oConect.CargaRecordSet(sql)
Exit Function
GetCajaOperacionesGenErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function



Public Function GetCajaGenConfirmacion(psOpecod As String, pdFechaDel As Date, pdFechaAl As Date, Optional pbVerReferencia As Boolean = False, Optional psOpeCodRef As String = "", Optional lnNroDigOperacion As Integer = 0) As ADODB.Recordset
Dim lsOpeFiltro As String
On Error GoTo GetCajaOperacionesGenErr
'lnNroDigOperacion = Nro de Caracteres del Codigo de Operación que se desea Buscar
If lnNroDigOperacion > 0 Then
    lsOpeFiltro = "LEFT(m.cOpeCod," & lnNroDigOperacion & ") = '" & Left(psOpecod, lnNroDigOperacion) & "' "
Else
    lsOpeFiltro = "m.cOpeCod = '" & psOpecod & "' "
End If
sql = "SELECT Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103) dFecha, ISNULL(d.cDocAbrev,''), ISNULL(md.cDocNro,'Ninguno'), Abs(" & IIf(Mid(psOpecod, 3, 1) = "2", "SUM( ISNULL(me.nMovMEImporte,0) )", "SUM(mc.nMovImporte)") & "), mg.cPersCod, p.cPersNombre, m.cMovDesc, m.cMovNro, m.nMovNro, ISNULL(md.nDocTpo,0) " _
    & "FROM Mov m inner join Movref mr on m.nmovnro=mr.nmovnro inner join Mov m1 on m1.nmovnro=mr.nmovnroref LEFT JOIN MovDoc md ON md.nMovNro = m1.nMovNro and not md.nDocTpo in (" & TpoDocVoucherEgreso & "," & TpoDocFormSUNAT & ") LEFT JOIN Documento d ON d.nDocTpo = md.nDocTpo LEFT JOIN MovCta mc ON mc.nMovNro = m1.nMovNro " _
    & IIf(Mid(psOpecod, 3, 1) = "2", " JOIN MovMe  me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
    & "      LEFT JOIN MovGasto mg ON mg.nMovNro = mc.nMovNro " _
    & "      LEFT JOIN Persona p ON p.cPersCod = mg.cPersCod " _
    & IIf(pbVerReferencia, " LEFT JOIN ( SELECT m1.nMovNro, mr.nMovNroRef FROM Mov m1 JOIN MovRef mr ON mr.nMovNro = m1.nMovNro WHERE m1.cOpecod = '" & psOpeCodRef & "' and m1.nMovEstado = " & gMovEstContabMovContable & " and m1.nMovFlag = " & gMovFlagVigente & " ) Ref ON Ref.nMovNroRef = m.nMovNro ", "") _
    & "WHERE " & lsOpeFiltro & "  and m.nMovFlag = " & gMovFlagVigente & " and mc.nmovimporte > 0 " _
    & "  and LEFT(m.cMovNro,8) BETWEEN '" & Format(pdFechaDel, gsFormatoMovFecha) & "' and '" & Format(pdFechaAl, gsFormatoMovFecha) & "' " _
    & IIf(pbVerReferencia, " and Ref.nMovNro IS NULL ", "") _
    & "GROUP BY Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103), " _
    & "      ISNULL(d.cDocAbrev,''), ISNULL(md.cDocNro,'Ninguno'), mg.cPersCod, p.cPersNombre, m.cMovDesc, m.cMovNro, m.nMovNro, ISNULL(md.nDocTpo,0) " _
    & " ORDER BY m.cMovNro "
Set GetCajaGenConfirmacion = oConect.CargaRecordSet(sql)
'and mc.nmovimporte > 0 and m.nMovEstado = " & gMovEstContabMovContable & "
Exit Function
GetCajaOperacionesGenErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function VerifBCAR(psPersCod As String) As Boolean
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

sql = " Select cPersBCRuc From PersSunatBC Where cPersBCRuc In" _
    & " (Select cPersIDNro From PersID Where cPersCod = '" & psPersCod & "' And cPersIDTpo  = " & PersIdTipo.gPersIdRUC & ")"
Set rs = oConect.CargaRecordSet(sql)

If rs.BOF And rs.EOF Then
    VerifBCAR = False
Else
    VerifBCAR = True
End If

rs.Close
Set rs = Nothing

Exit Function
GetCajaBancosOperacionesErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetMontoIngresoRetencion(psPersCod As String, psMovNro As String, pbIngresos As Boolean) As Currency
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

If pbIngresos Then
    sql = " Select dbo.RCFMontoPagado('" & psPersCod & "', '" & psMovNro & "')"
Else
    sql = " Select dbo.RCFMontoRetenido('" & psPersCod & "', '" & psMovNro & "')"
End If
Set rs = oConect.CargaRecordSet(sql)

If rs.BOF And rs.EOF Then
    GetMontoIngresoRetencion = 0
Else
    GetMontoIngresoRetencion = rs.Fields(0)
End If

rs.Close
Set rs = Nothing

Exit Function
GetCajaBancosOperacionesErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetCajaExtChequesRegSinDeposito(psOpecod As String, pdFechaDel As Date, pdFechaAl As Date) As ADODB.Recordset
On Error GoTo GetCajaExtChequesRegSinDepositoErr
sql = "SELECT Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103) dFecha, p.cPersNombre," _
    & "    dr.cIFCta,d .cDocAbrev, dr.cNroDoc, IsNull(dr.nMonto, 0), m.cMovDesc, dr.cPersCod, dr.cIFTpo, '' cCtaIFCod, m.cMovNro, m.nMovNro, md.nDocTpo, dr.cAreaCod, dr.cAgeCod, convert(varchar(10),md.dDocFecha,103) as FechaReg " _
    & "FROM Mov m JOIN MovDoc md ON md.nMovNro = m.nMovNro JOIN Documento d ON d.nDocTpo = md.nDocTpo " _
    & "           JOIN DocRec dr ON dr.cNroDoc = md.cDocNro and dr.nTpoDoc = md.nDocTpo " _
    & "           JOIN DocRecEst dre ON dre.cNroDoc = dr.cNroDoc and dre.nTpoDoc = dr.nTpoDoc and dre.cPersCod = dr.cPersCod and dre.cMovNro = m.cMovNro and dre.nEstado = 1 " _
    & "           JOIN Persona p ON p.cPersCod = dr.cPersCod " _
    & "WHERE m.cOpeCod = '" & psOpecod & "' and m.nMovEstado in (" & gMovEstContabMovContable & ", " & gMovEstContabNoContable & ") and m.nMovFlag = " & gMovFlagVigente & " and dr.nTpoDoc = " & TpoDocCheque & " " _
    & "  and LEFT(m.cMovNro,8) BETWEEN '" & Format(pdFechaDel, gsFormatoMovFecha) & "' and '" & Format(pdFechaAl, gsFormatoMovFecha) & "' "

Set GetCajaExtChequesRegSinDeposito = oConect.CargaRecordSet(sql)
Exit Function
GetCajaExtChequesRegSinDepositoErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetMovRefCartaFianza(ByVal pnMovNro As Double) As ADODB.Recordset
On Error GoTo GetMovRefCartaFianzaErr
    Dim sSql As String
    sSql = "select m.nmovnro,cmovnro from mov m inner join movref mr on m.nmovnro=mr.nmovnro where M.nMOVFLAG NOT  IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') and nmovnroref=" & pnMovNro
    
Set GetMovRefCartaFianza = oConect.CargaRecordSet(sSql)
Exit Function
GetMovRefCartaFianzaErr:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function




Public Function GrabaPagoProveedorNew(psMovNro As String, psOpecod As String, psMovDesc As String, psCtaContDebeB As String, psCtaContDebeS As String, _
                                   psCtaContHaber As String, pnImporteB As Currency, pnImporteS As Currency, psPersCod As String, psTpoIF As String, psPersCodIf As String, psCtaIf As String, _
                                   prsBilletaje As ADODB.Recordset, psDocTpo As String, psDocNro As String, psFechaDoc As String, psDocVoucher As String, rsDocs As ADODB.Recordset, psCuentaAho As String, _
                                   ByVal psCtaDiferencia As String, ByVal pnMontoDiferencia As Currency, Optional ByVal gbBitCentral As Boolean, Optional pbGrabaOpeNegocio As Boolean = False, Optional pnRetAct As Currency = 0, Optional lsCtaITFD As String, Optional lsCtaITFH As String, Optional nTasaITF As Double = 0, Optional bIncluyeNC As Boolean = False, _
                                   Optional bAfectoITF As Boolean = True, Optional lnITFDescCuentas As Double) As Integer
Dim lnMovItem  As Long
Dim lnMovOrden As Integer
Dim lnMovNro As Long
Dim lbTrans  As Boolean
Dim lsMovNro As String
Dim lsCadBol As String
Dim oDMov    As DMov
Set oDMov = New DMov

Dim oConst As NConstSistemas
Set oConst = New NConstSistemas
Dim oImp As DImpuesto
Set oImp = New DImpuesto
Dim oFun As NContFunciones
Set oFun = New NContFunciones
Dim lbBitReten As Boolean
Dim lsCtaReten As String
Dim lsCtaRetenConvesion As String
Dim lsCtaRetenCambio As String
Dim lbBCAR As Boolean
Dim lnTasaImp As Currency
Dim lnIngresos As Currency
Dim lnRetencion As Currency
Dim lnTopeRetencion As Currency
Dim lnRetAct As Currency
Dim lnRetActME As Currency
Dim lsRetDocNro As String
Dim lMN As Boolean
Dim lsMovNroNegocio As String
Dim lbGraboNegocio As Boolean

Dim nTCFijo As Currency
Dim nTCVenta As Currency
Dim nTCComp As Currency
Dim dFecha As String

On Error GoTo GrabaPagoProveedorErr

GetTipCambio GetFechaMov(psMovNro, True)

If Mid(psCtaContHaber, 3, 1) = Moneda.gMonedaExtranjera Then
    lMN = False
Else
    lMN = True
End If

lbBitReten = IIf(oConst.LeeConstSistema(gConstSistBitRetencion6Porcent) = 1, True, False)
If lbBitReten And pnRetAct <> 0 Then
   lsRetDocNro = oFun.GeneraDocNro(TpoDocAlmacenComprdeRetencion6)
    lsCtaReten = oConst.LeeConstSistema(gConstSistCtaRetencion6Porcent)
End If

If psOpecod <> "402581" Then
    If Mid(psOpecod, 3, 1) = 2 Then
       dFecha = CDate(psFechaDoc)
       oDMov.RecuperaTipoCambioF dFecha, psOpecod, , , nTCFijo, nTCVenta, nTCComp
    End If
End If



oDMov.BeginTrans
lbTrans = True
GrabaPagoProveedorNew = 1

oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporteB + pnImporteS, "0", "0"

'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
If pnImporteB <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCtaPer lnMovNro, lnMovItem, psCtaContDebeB, pnImporteB
End If
If pnImporteS <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCtaPer lnMovNro, lnMovItem, psCtaContDebeS, pnImporteS
End If

'Agregado

lnRetAct = 0
lnRetActME = 0
If lbBitReten Then
   lnRetAct = pnRetAct
   ''   lnRetActME = Round(lnRetAct / gnTipCambioPonderado, 2)
   lnRetActME = Round(lnRetAct / gnTipCambioPonderado, 2)
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCtaPer lnMovNro, lnMovItem, lsCtaReten, lnRetAct * -1
   oDMov.InsertaMovDoc lnMovNro, TpoDocAlmacenComprdeRetencion6, lsRetDocNro, psFechaDoc
End If

'guardamos el billetaje si este posee
If Not prsBilletaje Is Nothing Then
   oDMov.GrabaMovimientoEfectivoNew lnMovNro, lnMovItem, prsBilletaje, psCtaContHaber, "H"
Else
    'guardamos la cuenta del haber
'    If (pnImporteB + pnImporteS - lnRetAct) <> 0 Then
    If (pnImporteB + pnImporteS - IIf(lMN, lnRetAct, lnRetActME)) <> 0 Then
        lnMovItem = lnMovItem + 1: lnMovOrden = 0
'        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber, (pnImporteB + pnImporteS - IIf(lnRetAct > 0, lnRetAct, 0)) * -1
        oDMov.InsertaMovCtaPer lnMovNro, lnMovItem, psCtaContHaber, (pnImporteB + pnImporteS - IIf(lnRetAct + lnRetActME > 0, IIf(lMN, lnRetAct, lnRetActME), 0)) * -1
    End If
    If psPersCodIf <> "" Then
        'guardamos la Cta de la IF si es que utiliza esta
        lnMovOrden = lnMovOrden + 1
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCodIf, psTpoIF, psCtaIf
        
    End If
End If
If psDocNro <> "" Then
   oDMov.InsertaMovDoc lnMovNro, psDocTpo, psDocNro, psFechaDoc
End If
If psDocVoucher <> "" Then
   oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocVoucher, GetFechaMov(psMovNro, False)
End If


oDMov.InsertaMovGasto lnMovNro, psPersCod, "0"

If pnMontoDiferencia <> 0 Then
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCtaPer lnMovNro, Format(lnMovItem, "#0"), psCtaDiferencia, pnMontoDiferencia * -1
End If

'Referenciamos todos los Documentos que se Pagan
If bIncluyeNC Then
    '   OJO: Pago Total
    If Not rsDocs Is Nothing Then
       Do While Not rsDocs.EOF
          If rsDocs!Itm = "1" Then
             oDMov.InsertaMovRef lnMovNro, rsDocs!nMovNro
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeB, 0, False
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeS, 0, False
         
          End If
          rsDocs.MoveNext
       Loop
    End If
Else
    ' Puede ser pago parcial
    
    Dim nMontoPago As Currency
    If Not rsDocs Is Nothing Then
        nMontoPago = pnImporteB + pnImporteS
       Do While Not rsDocs.EOF And nMontoPago > 0
          If rsDocs!Itm = "1" Then
             oDMov.InsertaMovRef lnMovNro, rsDocs!nMovNro
             If nMontoPago >= rsDocs!Importe Then
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeB, 0, False
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeS, 0, False
                nMontoPago = nMontoPago - rsDocs!Importe
             Else
                If pnImporteB <> 0 Then
                    oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebeB, rsDocs!Importe - nMontoPago
                Else
                    oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebeS, rsDocs!Importe - nMontoPago
                End If
                nMontoPago = 0
             End If
          End If
          rsDocs.MoveNext
       Loop
    End If

End If

'Fin Agregado

'--Calculo del ITF
'--  JOHN
If bAfectoITF = True Then 'bit si esta afecto al ITF el pago a Proveedor
    'Dim lnMontoITF As Currency
    Dim lnMontoITF As Double '*** PEAC 20110331
    If lsCtaITFD <> "" And lsCtaITFH <> "" And nTasaITF > 0 Then
        lnMontoITF = fgTruncar((pnImporteB + pnImporteS - lnRetAct) * nTasaITF, 2)
        If lnMontoITF > 0 Then
            lnMovItem = lnMovItem + 1
            oDMov.InsertaMovCtaPer lnMovNro, Format(lnMovItem, "#0"), lsCtaITFD, lnMontoITF
            lnMovItem = lnMovItem + 1
            oDMov.InsertaMovCtaPer lnMovNro, Format(lnMovItem, "#0"), lsCtaITFH, lnMontoITF * -1
        End If
    End If
End If
'----------

If psOpecod <> "402581" Then
    If Mid(psOpecod, 3, 1) = 2 Then
        If lbBitReten Then
            oDMov.GeneraMovMEPerf lnMovNro, psMovNro, gnTipCambioPonderado, nTCFijo, nTCVenta, nTCComp
            'oDMov.GeneraMovME lnMovNro, psMovNro, gnTipCambioPonderado
        Else
            oDMov.GeneraMovMEPerf lnMovNro, psMovNro, , nTCFijo, nTCVenta, nTCComp
            'oDMov.GeneraMovME lnMovNro, psMovNro
        End If
    End If
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"

If psDocTpo = TpoDocNotaAbono Or psDocTpo = TpoDocNotaCargo Then
    oDMov.InsertaRegDocCuenta psDocTpo, psDocNro, psCuentaAho
    lbGraboNegocio = False
    If pbGrabaOpeNegocio Then
        Dim oCapta As NCapMovimientos
        Dim lnSaldo As Double
        Set oCapta = New NCapMovimientos
        Dim oCon As New DConecta
        Dim oDis As New NRHProcesosCierre
        If gbBitCentral Then
           lsMovNroNegocio = oDMov.GeneraMovNro(, , , psMovNro)
           lbGraboNegocio = True
           lnSaldo = oCapta.CapAbonoCuentaAho(psCuentaAho, pnImporteB + pnImporteS - lnRetAct, gCGPagProvAbonoCent, lsMovNroNegocio, psMovDesc, TpoDocNotaAbono, psDocNro, , , , , , , , , , True, False, , , oDMov.GetConexion, False, bAfectoITF, CDbl(lnITFDescCuentas))
        Else
            If oCon.AbreConexion(Left(psCuentaAho, 2)) Then
                oCon.BeginTrans
                lbGraboNegocio = True
                lnSaldo = oDis.Abono(psCuentaAho, pnImporteB + pnImporteS - lnRetAct, gCGPagProvAbonoDist, gCGPagProvAbonoDist, "112" & Left(psCuentaAho, 2), Right(psMovNro, 4), psDocNro, "PAGO : " & psMovDesc, oCon, GetFechaMov(psMovNro, True))
                oCon.CommitTrans
            End If
            oCon.CierraConexion
            Set oCon = Nothing
            Set oDis = Nothing
        End If
    End If
End If

oDMov.CommitTrans
lbTrans = False
GrabaPagoProveedorNew = 0


Set oDMov = Nothing

Exit Function
GrabaPagoProveedorErr:
   lsMsgErr = Err.Description
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    If lbGraboNegocio Then
        lsMsgErr = lsMsgErr & Chr(10) & Chr(10) & "EL SISTEMA HIZO EL ABONO EN LA CUENTA DEL CLIENTE. POR FAVOR EXTORNAR ESTA OPERACION"
    End If
    Set oDMov = Nothing
    Err.Raise vbObjectError + 100, "GrabaPagoProveedor", lsMsgErr
End Function

Public Function GrabaPagoProveedorDetraccion(psMovNro As String, psOpecod As String, psMovDesc As String, psCtaContDebeB As String, psCtaContDebeS As String, _
                                   psCtaContHaber As String, pnImporteB As Currency, pnImporteS As Currency, psPersCod As String, psTpoIF As String, psPersCodIf As String, psCtaIf As String, _
                                   prsBilletaje As ADODB.Recordset, psDocTpo As String, psDocNro As String, psFechaDoc As String, psDocVoucher As String, rsDocs As ADODB.Recordset, psCuentaAho As String, _
                                   ByVal psCtaDiferencia As String, ByVal pnMontoDiferencia As Currency, Optional ByVal gbBitCentral As Boolean, Optional pbGrabaOpeNegocio As Boolean = False, Optional pnRetAct As Currency = 0, Optional lsCtaITFD As String, Optional lsCtaITFH As String, Optional nTasaITF As Double = 0, Optional bIncluyeNC As Boolean = False, _
                                   Optional bAfectoITF As Boolean = True, Optional lnITFDescCuentas As Double, Optional pnImporteRH As Currency, Optional psCtaContDebeRH As String, Optional psCtacontDebeRHJ As String, Optional pnImporteRHJ As Currency, Optional psCtaContDebeSegu As String, Optional pnImporteSegu As Currency, Optional psFechaDocRef As String, Optional psDocRefNro As String) As Integer
Dim lnMovItem  As Long
Dim lnMovOrden As Integer
Dim lnMovNro As Long
Dim lbTrans  As Boolean
Dim lsMovNro As String
Dim lsCadBol As String
Dim oDMov    As DMov
Set oDMov = New DMov

Dim oConst As NConstSistemas
Set oConst = New NConstSistemas
Dim oImp As DImpuesto
Set oImp = New DImpuesto
Dim oFun As NContFunciones
Set oFun = New NContFunciones
Dim lbBitReten As Boolean
Dim lsCtaReten As String
Dim lsCtaRetenConvesion As String
Dim lsCtaRetenCambio As String
Dim lbBCAR As Boolean
Dim lnTasaImp As Currency
Dim lnIngresos As Currency
Dim lnRetencion As Currency
Dim lnTopeRetencion As Currency
Dim lnRetAct As Currency
Dim lnRetActME As Currency
Dim lsRetDocNro As String
Dim lMN As Boolean
Dim lsMovNroNegocio As String
Dim lbGraboNegocio As Boolean

Dim lsTextError As String '*** PEAC 20111115

On Error GoTo GrabaPagoProveedorErr

GetTipCambio GetFechaMov(psMovNro, True)

If Mid(psCtaContHaber, 3, 1) = Moneda.gMonedaExtranjera Then
    lMN = False
Else
    lMN = True
End If

lbBitReten = IIf(oConst.LeeConstSistema(gConstSistBitRetencion6Porcent) = 1, True, False)
If lbBitReten And pnRetAct <> 0 Then
   lsRetDocNro = oFun.GeneraDocNro(TpoDocAlmacenComprdeRetencion6)
    lsCtaReten = oConst.LeeConstSistema(gConstSistCtaRetencion6Porcent)
End If

oDMov.BeginTrans
lbTrans = True
GrabaPagoProveedorDetraccion = 1

oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu, "0", "0"

'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
If pnImporteB <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeB, pnImporteB
End If
If pnImporteS <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeS, pnImporteS
End If
'JEOM
If pnImporteRH <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeRH, pnImporteRH
End If
If pnImporteRHJ <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtacontDebeRHJ, pnImporteRHJ
End If
If pnImporteSegu <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeSegu, pnImporteSegu
End If

'Agregado

lnRetAct = 0
lnRetActME = 0
If lbBitReten Then
   lnRetAct = pnRetAct
   ''   lnRetActME = Round(lnRetAct / gnTipCambioPonderado, 2)
   lnRetActME = Round(lnRetAct / gnTipCambioPonderado, 2)
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCtaReten, lnRetAct * -1
   oDMov.InsertaMovDoc lnMovNro, TpoDocAlmacenComprdeRetencion6, lsRetDocNro, psFechaDoc
End If

'guardamos el billetaje si este posee
If Not prsBilletaje Is Nothing Then
   oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, prsBilletaje, psCtaContHaber, "H"
Else
    'guardamos la cuenta del haber
'    If (pnImporteB + pnImporteS - lnRetAct) <> 0 Then
    If (pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu - IIf(lMN, lnRetAct, lnRetActME)) <> 0 Then
        lnMovItem = lnMovItem + 1: lnMovOrden = 0
'        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber, (pnImporteB + pnImporteS - IIf(lnRetAct > 0, lnRetAct, 0)) * -1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber, (pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu - IIf(lnRetAct + lnRetActME > 0, IIf(lMN, lnRetAct, lnRetActME), 0)) * -1, , lsTextError
        
        '*** PEAC 20111115
        If Len(lsTextError) > 0 Then
            
        End If
        '*** FIN PEAC
        
    End If
    If psPersCodIf <> "" Then
        'guardamos la Cta de la IF si es que utiliza esta
        lnMovOrden = lnMovOrden + 1
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCodIf, psTpoIF, psCtaIf
        
    End If
End If
If psDocNro <> "" Then
   oDMov.InsertaMovDoc lnMovNro, psDocTpo, psDocNro, psFechaDoc
End If
If psDocVoucher <> "" Then
   oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocVoucher, GetFechaMov(psMovNro, False)
End If

'JEOM Documento de Referencia para Detraccion
If psDocRefNro <> "" And psFechaDocRef <> "" Then
   oDMov.InsertaMovDocRef lnMovNro, psDocRefNro, Format(CDate(psFechaDocRef), gsFormatoFecha), 2   ' Tipo 2 para Detraccion
End If


oDMov.InsertaMovGasto lnMovNro, psPersCod, "0"

If pnMontoDiferencia <> 0 Then
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDiferencia, pnMontoDiferencia * -1
End If

'Referenciamos todos los Documentos que se Pagan
If bIncluyeNC Then
    '   OJO: Pago Total
    If Not rsDocs Is Nothing Then
       Do While Not rsDocs.EOF
          If rsDocs!Itm = "1" Then
             oDMov.InsertaMovRef lnMovNro, rsDocs!nMovNro
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeB, 0, False
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeS, 0, False
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeRH, 0, False
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtacontDebeRHJ, 0, False
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeSegu, 0, False
         
          End If
          rsDocs.MoveNext
       Loop
    End If
Else
    ' Puede ser pago parcial
    
    Dim nMontoPago As Currency
    If Not rsDocs Is Nothing Then
        nMontoPago = pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu
       Do While Not rsDocs.EOF And nMontoPago > 0
          If rsDocs!Itm = "1" Then
             oDMov.InsertaMovRef lnMovNro, rsDocs!nMovNro
             If nMontoPago >= rsDocs!Importe Then
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeB, 0, False
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeS, 0, False
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeRH, 0, False
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtacontDebeRHJ, 0, False
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeSegu, 0, False
                nMontoPago = nMontoPago - rsDocs!Importe
             Else
                If pnImporteB <> 0 Then
                    oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebeB, rsDocs!Importe - nMontoPago
                Else
                   If pnImporteS <> 0 Then
                      oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebeS, rsDocs!Importe - nMontoPago
                   Else
                      If pnImporteRH <> 0 Then
                        oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebeRH, rsDocs!Importe - nMontoPago
                      Else
                        If pnImporteRHJ <> 0 Then
                           oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtacontDebeRHJ, rsDocs!Importe - nMontoPago
                        Else
                            oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebeSegu, rsDocs!Importe - nMontoPago
                        End If
                      End If
                   End If
                End If
                nMontoPago = 0
             End If
          End If
          rsDocs.MoveNext
       Loop
    End If

End If

'Fin Agregado

'--Calculo del ITF
'--  JOHN
If bAfectoITF = True Then 'bit si esta afecto al ITF el pago a Proveedor
    'Dim lnMontoITF As Currency
    Dim lnMontoITF As Double '*** PEAC 20110331
    If lsCtaITFD <> "" And lsCtaITFH <> "" And nTasaITF > 0 Then
        lnMontoITF = fgTruncar((pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu - lnRetAct) * nTasaITF, 2)
        If lnMontoITF > 0 Then
            lnMovItem = lnMovItem + 1
            oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFD, lnMontoITF
            lnMovItem = lnMovItem + 1
            oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFH, lnMontoITF * -1
        End If
    End If
End If
'----------

If psOpecod <> "402581" And psOpecod <> "402582" Then
    If Mid(psOpecod, 3, 1) = 2 Then
        If lbBitReten Then
            oDMov.GeneraMovME lnMovNro, psMovNro, gnTipCambioPonderado
        Else
            oDMov.GeneraMovME lnMovNro, psMovNro
        End If
    End If
End If

oDMov.ActualizaSaldoMovimiento psMovNro, "+"

If psDocTpo = TpoDocNotaAbono Or psDocTpo = TpoDocNotaCargo Then
    oDMov.InsertaRegDocCuenta psDocTpo, psDocNro, psCuentaAho
    lbGraboNegocio = False
    If pbGrabaOpeNegocio Then
        Dim oCapta As NCapMovimientos
        Dim lnSaldo As Double
        Set oCapta = New NCapMovimientos
        Dim oCon As New DConecta
        Dim oDis As New NRHProcesosCierre
        If gbBitCentral Then
           lsMovNroNegocio = oDMov.GeneraMovNro(, , , psMovNro)
           lbGraboNegocio = True
           lnSaldo = oCapta.CapAbonoCuentaAho(psCuentaAho, pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu - lnRetAct, gCGPagProvAbonoCent, lsMovNroNegocio, psMovDesc, TpoDocNotaAbono, psDocNro, , , , , , , , , , True, False, , , oDMov.GetConexion, False, bAfectoITF, CDbl(lnITFDescCuentas))
        Else
            If oCon.AbreConexion(Left(psCuentaAho, 2)) Then
                oCon.BeginTrans
                lbGraboNegocio = True
                lnSaldo = oDis.Abono(psCuentaAho, pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu - lnRetAct, gCGPagProvAbonoDist, gCGPagProvAbonoDist, "112" & Left(psCuentaAho, 2), Right(psMovNro, 4), psDocNro, "PAGO : " & psMovDesc, oCon, GetFechaMov(psMovNro, True))
                oCon.CommitTrans
            End If
            oCon.CierraConexion
            Set oCon = Nothing
            Set oDis = Nothing
        End If
    End If
End If

oDMov.CommitTrans
lbTrans = False
GrabaPagoProveedorDetraccion = 0

Set oDMov = Nothing

Exit Function
GrabaPagoProveedorErr:
   lsMsgErr = Err.Description
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    If lbGraboNegocio Then
        lsMsgErr = lsMsgErr & Chr(10) & Chr(10) & "EL SISTEMA HIZO EL ABONO EN LA CUENTA DEL CLIENTE. POR FAVOR EXTORNAR ESTA OPERACION"
    End If
    Set oDMov = Nothing
    Err.Raise vbObjectError + 100, "GrabaPagoProveedor", lsMsgErr
End Function

Public Function GrabaDeposiBancoNegocioFinan(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
                                        ByVal psCtaContBanco As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                        ByVal pnDocTpo As TpoDoc, ByVal psDocNro As String, ByVal psFechaDoc As String, _
                                        ByVal psDocNroVoucher As String, _
                                        ByVal lnTpoObjIf As TpoObjetos, ByVal psCtaIf As String, ByVal pbDeposito As Boolean, _
                                        Optional pnMotivoNANC As Long = -1, Optional psMotObjPadre As String = "", Optional psMotObjeto As String = "", Optional ByVal psctaAho As String = "", _
                                        Optional ByVal pnDocTpoNANCAux As TpoDoc = -1, Optional ByVal psNroDocNANCAux As String = "", Optional ByVal pnImporteAux As Currency, _
                                        Optional pnMotivoNANCAux As Long = -1, Optional psMotObjPadreAux As String = "", Optional psMotObjetoAux As String = "", Optional ByVal psCtaAhoAux As String = "", _
                                        Optional ByVal pbGrabaNegocio As Boolean = False, Optional ByVal pbBitCentral As Boolean = False, _
                                        Optional psOpeCodNegocio As String = "", _
                                        Optional psAjusteCta As String = "", Optional pnAjusteMonto As Currency = 0, Optional pnTCBanco As Currency = 0, Optional pnMovNroRef As Long) As Integer
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
Dim lsMovNroNegocio As String
Set oCont = New NContFunciones
On Error GoTo ErrorGrabaDepRetBancosVarios

GrabaDeposiBancoNegocioFinan = 1
oDMov.BeginTrans
lbTrans = True

oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1
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
        Case Else
            lnMovOrden = lnMovOrden + 1
            oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, psCtaIf
    End Select
End If


lnMovItem = lnMovItem + 1: lnMovOrden = 0

oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber + lsSubCta, (pnImporte - pnAjusteMonto) * IIf(pbDeposito, -1, 1)

If psCtaHaber = "111109" Or psCtaHaber = "112109" Then
    
    oDMov.InsertaMovHabilitaCaja lnMovNro, Abs((pnImporte - pnAjusteMonto)), "025", "026", lsSubCta, Mid(psCtaHaber, 3, 1)
    
    If psCtaHaber = "111109" Then
        oDMov.InsertaMovObjEfectivo lnMovNro, 1, 1, "1M005"
    Else
        oDMov.InsertaMovObjEfectivo lnMovNro, 1, 1, "2B005"
    End If
End If

'If Not rsObjHaber Is Nothing Then
'    If rsObjHaber.State = adStateOpen Then
'        rsObjHaber.MoveFirst
'        Do While Not rsObjHaber.EOF
'            lnMovOrden = lnMovOrden + 1
'            Select Case rsObjHaber!cObjetoCod
'                Case ObjCMACAgenciaArea
'                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!cObjetoCod
'                    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(rsObjHaber!objeto, 4, 2), Mid(rsObjHaber!objeto, 1, 3)
'                Case ObjCMACAgencias
'                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!cObjetoCod
'                    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!objeto, "013"
'                Case ObjDescomEfectivo
'                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!cObjetoCod
'                    oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!objeto
'                Case ObjEntidadesFinancieras
'                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!cObjetoCod
'                    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(rsObjHaber!objeto, 4, 13), Mid(rsObjHaber!objeto, 1, 2), Mid(rsObjHaber!objeto, 18, 10)
'                Case Else
'                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!objeto
'            End Select
'            rsObjHaber.MoveNext
'        Loop
'    End If
'End If

If pnAjusteMonto <> 0 Then
   oDMov.InsertaMovCta lnMovNro, lnMovItem + 1, psAjusteCta, pnAjusteMonto * IIf(pbDeposito, -1, 1)
End If

Select Case Val(pnDocTpo)
    Case TpoDocNotaAbono, TpoDocNotaCargo
        oDMov.InsertaNotaAbonoCargo pnDocTpo, psDocNro, gNCNAEnMovimiento, 0, pnImporte, psCtaIf, ObjEntidadesFinancieras, Mid(psOpecod, 3, 1)
'        oDMov.ActualizaNotaAbonoCargo pnDocTpo, psDocNro, gNCNAEnMovimiento
End Select

Select Case Val(pnDocTpo)
    Case TpoDocCarta, TpoDocCheque, TpoDocNotaCargo
        If pbGrabaNegocio Then
            Dim oCapta As NCapMovimientos
            Dim lnSaldo As Double
            Set oCapta = New NCapMovimientos
            Dim oCon As New DConecta
            Dim oDis As New NRHProcesosCierre
            
            If pbBitCentral Then
               lsMovNroNegocio = oDMov.GeneraMovNro(, , , psMovNro)
               lnSaldo = oCapta.CapCargoCuentaAho(psCtaAhoAux, pnImporteAux, psOpeCodNegocio, lsMovNroNegocio, psMovDesc, TpoDocNotaCargo, psDocNro, Mid(psCtaIf, 4, 13), , , , , , , , , True, "", , oDMov.GetConexion, False)
            Else
                If oCon.AbreConexion Then 'Remota(Left(psCtaAhoAux, 2))
                    oCon.BeginTrans
                    lnSaldo = oDis.Cargo(psCtaAhoAux, pnImporteAux, psOpeCodNegocio, "112" & Left(psCtaAhoAux, 2), Right(psMovNro, 4), psDocNro, "CARGO : " & psMovDesc, oCon, GetFechaMov(psMovNro, True))
                    oCon.CommitTrans
                End If
                oCon.CierraConexion
                Set oCon = Nothing
                Set oDis = Nothing
            End If
        End If
End Select

If psDocNro <> "" Then
    oDMov.InsertaMovDoc lnMovNro, pnDocTpo, psDocNro, Format(CDate(psFechaDoc), gsFormatoFecha)
End If
If psDocNroVoucher <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocNroVoucher, Format(CDate(psFechaDoc), gsFormatoFecha)
End If
If pnDocTpoNANCAux <> -1 And psNroDocNANCAux <> "" Then
    oDMov.InsertaNotaAbonoCargo pnDocTpoNANCAux, psNroDocNANCAux, gNCNAEnMovimiento, 0, pnImporteAux, psCtaIf, ObjEntidadesFinancieras, Mid(psOpecod, 3, 1)
    oDMov.InsertaMovDoc lnMovNro, pnDocTpoNANCAux, psNroDocNANCAux, Format(CDate(psFechaDoc), gsFormatoFecha)
End If

Dim lsCtaCaracterOut As String
If oDMov.CuentaEsPendiente(psCtaHaber + lsSubCta, lsCtaCaracterOut, IIf(pbDeposito, "A", "D")) Then
    If (lsCtaCaracterOut = "A" And pbDeposito) Or (lsCtaCaracterOut = "D" And Not pbDeposito) Then
        oDMov.InsertaMovPendientesRend lnMovNro, psCtaHaber + lsSubCta, pnImporte
    End If
End If

oDMov.InsertaMovRef lnMovNro, pnMovNroRef

GrabaDeposiBancoNegocioFinan = 0
If Mid(psOpecod, 3, 1) = 2 Then
    oDMov.GeneraMovME lnMovNro, psMovNro, pnTCBanco
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"


oDMov.CommitTrans
lbTrans = False
Set oDMov = Nothing
Exit Function
ErrorGrabaDepRetBancosVarios:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaRendicionGiroDocumento", lsMsgErr
End Function
'JEOM
Public Function GrabaDepositoRegularizacion(ByVal psMovNro As String, ByVal pnMovNroRef As Long, ByVal psOpecod As String, ByVal psMovDesc As String, ByVal pnImporte As Currency) As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lnMovNro As Long
Dim lbTrans As Boolean

Set oDMov = New DMov
On Error GoTo ErrorGrabaDepRetBancosVarios

GrabaDepositoRegularizacion = 1
oDMov.BeginTrans
lbTrans = True

oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
oDMov.InsertaMovRef lnMovNro, pnMovNroRef


oDMov.CommitTrans
lbTrans = False
Set oDMov = Nothing
Exit Function
ErrorGrabaDepRetBancosVarios:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaRendicionGiroDocumento", lsMsgErr
End Function

'JACA 20110908********************************************************

Public Function guardarInversionApertura(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
       ByVal prsCtaApe As ADODB.Recordset, pbCreaSubCta As Boolean, psSubCtaIF As String, ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, _
       ByVal pnMoneda As Integer, ByVal pdFecsis As Date, ByVal psEntTransferIFCod As String, ByVal pnTpoInversion As Integer, _
       pnCondicion As Integer) As Integer
            
Dim oDMov      As DMov
Dim oOpe       As DOperacion

Dim lnMovItem  As Long
Dim lnMovOrden As Long
Dim lnMovNro   As Long
Dim lbTrans    As Boolean
Dim lsCuenta   As String

On Error GoTo GrabaAperturaCtaIFErr

Set oDMov = New DMov
Set oOpe = New DOperacion

guardarInversionApertura = 1

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, prsCtaApe!Importe, "0", "0"

'Guardamos la cuenta en el debe
    lsCuenta = oOpe.EmiteOpeCta(psOpecod, "D", "0", Format(psIFTpo, "00") & psPersCod & psSubCtaIF, ObjEntidadesFinancieras)
    
    If pbCreaSubCta Then
        oDMov.InsertaCuentaIF psPersCod, psIFTpo, psSubCtaIF, "REGISTRO INVERSION " + IIf(pnMoneda = 1, "MN", "ME"), prsCtaApe!Fec_Apertura, Format(prsCtaApe!Fec_Apertura, gsFormatoFecha), Format(prsCtaApe!Fec_Apertura, gsFormatoFecha), 0, prsCtaApe!Fec_Apertura, gEstadoCtaIFActiva, psMovNro
    End If
    lnMovItem = 0
    
    lnMovItem = lnMovItem + 1: lnMovOrden = 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta, prsCtaApe!Importe
    If pnTpoInversion = 3 Then 'Si es Fonfo Mutuo
        oDMov.InsertaCuentaIF psPersCod, psIFTpo, prsCtaApe!Codigo, prsCtaApe!Cuenta_Nro, prsCtaApe!Fec_Apertura, "", "", 0, prsCtaApe!Fec_Apertura, _
                             gEstadoCtaIFRegistrada, psMovNro
    Else
        oDMov.InsertaCuentaIF psPersCod, psIFTpo, prsCtaApe!Codigo, prsCtaApe!Cuenta_Nro, prsCtaApe!Fec_Apertura, "", "", nVal(prsCtaApe!Plazo), prsCtaApe!Fec_Venc, _
                             gEstadoCtaIFRegistrada, psMovNro
    End If
    If pnTpoInversion <> 3 Then 'Si NO es Fonfo Mutuo
        oDMov.InsertaCuentaIFInteres psPersCod, psIFTpo, prsCtaApe!Codigo, prsCtaApe!Fec_Apertura, nVal(prsCtaApe!Plazo), nVal(prsCtaApe![Tasa]), psMovNro
    End If
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, Format(psIFTpo, "00"), prsCtaApe!Codigo
    
'Guardamos la cuenta en el Haber
'Transferencia
    Dim sOrden As String
    If Val(Mid(psEntTransferIFCod, 1, 2)) = gTpoIFCmac Then
        sOrden = "1"
    Else
        sOrden = "0"
    End If
    lnMovItem = lnMovItem + 1: lnMovOrden = 0
    
    lsCuenta = oOpe.EmiteOpeCta(psOpecod, "H", sOrden, psEntTransferIFCod, ObjEntidadesFinancieras)
    
    If lsCuenta = "" Then
        Err.Raise 50001, "nCajaGeneral: guardarInversionApertura", "Cuenta Contable de Inst.Financ Origen vacia. Consultar con Sistemas"
    End If
    
    oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta, prsCtaApe!Importe * -1
    lnMovOrden = lnMovOrden + 1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psEntTransferIFCod, 4, 13), Mid(psEntTransferIFCod, 1, 2), Mid(psEntTransferIFCod, 18, 10)
    
    If pnTpoInversion = 3 Then 'Si es Fonfo Mutuo
        oDMov.InsertaInversion lnMovNro, psOpecod, psPersCod, Format(psIFTpo, "00"), prsCtaApe!Codigo, prsCtaApe!Cuenta_Nro, _
                                   pnTpoInversion, pnMoneda, pnCondicion, prsCtaApe!Importe, 0, _
                                   prsCtaApe!Fec_Apertura, prsCtaApe!Fec_Apertura, prsCtaApe!Fec_Apertura, 0, 0, _
                                   Mid(psEntTransferIFCod, 4, 13), Mid(psEntTransferIFCod, 1, 2), Mid(psEntTransferIFCod, 18, 10), _
                                   psMovDesc, psMovNro
        
        oDMov.InsertaInversionFondosMutuos lnMovNro, prsCtaApe!Fec_Apertura, psPersCod, Format(psIFTpo, "00"), prsCtaApe!Codigo, prsCtaApe!Importe, prsCtaApe!Val_Cuota_1, prsCtaApe!Nro_Cuotas, 0
    Else
        oDMov.InsertaInversion lnMovNro, psOpecod, psPersCod, Format(psIFTpo, "00"), prsCtaApe!Codigo, prsCtaApe!Cuenta_Nro, _
                                   pnTpoInversion, pnMoneda, pnCondicion, prsCtaApe!Importe, prsCtaApe!Plazo, _
                                   prsCtaApe!Fec_Apertura, prsCtaApe!Fec_Venc, prsCtaApe!Fec_Apertura, prsCtaApe!Tasa, IIf(prsCtaApe!Int_Pactado = "", 0#, prsCtaApe!Int_Pactado), _
                                   Mid(psEntTransferIFCod, 4, 13), Mid(psEntTransferIFCod, 1, 2), Mid(psEntTransferIFCod, 18, 10), _
                                   psMovDesc, psMovNro
    
    End If
                                   
If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If


oDMov.ActualizaSaldoMovimiento psMovNro, "+"

oDMov.CommitTrans
lbTrans = False
guardarInversionApertura = 0
Set oDMov = Nothing
'Set oContF = Nothing

Exit Function
GrabaAperturaCtaIFErr:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "guardarInversionApertura", lsMsgErr & " Error Graba Apertura de Inversiones"
End Function

Public Function obtenerInversionesApertura(psOpecod As String, pdFechaIni As Date, pdFechaFin As Date, psTipoInversion As String) As Recordset
    On Error GoTo MsgError
    
    Dim rs As New ADODB.Recordset
    
    sql = "exec stp_sel_InversionApertura '" & psOpecod & "','" & Format(pdFechaIni, "yyyymmdd") & "','" & Format(pdFechaFin, "yyyymmdd") & "','" & psTipoInversion & "'"
    
    Set rs = oConect.CargaRecordSet(sql)
    Set obtenerInversionesApertura = rs
        
    Exit Function
MsgError:
    lsMsgErr = Err.Description
    Err.Raise vbObjectError + 100, "obtenerInversionesApertura", lsMsgErr & " Error Obtener Inversiones Aperturados"
End Function
Public Function getListaCtaContInversiones(psMoneda As String, ByVal pnTpoInv As Integer) As Recordset
'pnTpoInv Agregado xPASI20150921 ERS0712015
    On Error GoTo MsgError
    Dim sCtaFiltro As String
    Dim rs As New ADODB.Recordset

    sql = "exec stp_sel_ListaCtaContInversiones '" & psMoneda & "'," & pnTpoInv
    Set rs = oConect.CargaRecordSet(sql)

    Set getListaCtaContInversiones = rs
    
    Exit Function
MsgError:
    lsMsgErr = Err.Description
    Err.Raise vbObjectError + 100, "getListaCtaContInversiones", lsMsgErr & " Error Obtener Lista Cuentas Contables Inversiones"
End Function
Public Function GrabaConfInversion(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
                                   ByVal psCtaDebe As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                   ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psCtaIFCod As String, _
                                   ByVal pnMovNroRef As Long, ByVal pdFecha As Date, _
                                   ByVal pnTpoInversion As Integer, ByVal pnMoneda As Integer, ByVal pnCondicion As Integer, _
                                   ByVal pnPlazo As Integer, ByVal pdFecApe As Date, ByVal pdFecVenc As Date, _
                                   ByVal pnTasa As Double, ByVal pnIntPactado As Currency, _
                                   ByVal psPerCodTrans As String, ByVal psIFTipoTrans As String, ByVal psCtaIfCodTrans As String, _
                                   ByVal psCtaIFDesc As String, Optional ByVal psMovNroACAdd As String = "") As Integer
                                   'agrego psMovNroACAdd PASI20150921 ERS0712015
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lnMovNroACAdd As Long 'PASI20150921 ERS0712015
'Dim oCtaIf As NCajaCtaIF

Dim ldFecha As Date
Dim lsCuenta As String

'Set oCtaIf = New NCajaCtaIF
Set oDMov = New DMov

On Error GoTo MsgError



GrabaConfInversion = 1
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpecod, psMovDesc
lnMovNro = oDMov.GetnMovNro(psMovNro)
ldFecha = CDate(GetFechaMov(psMovNro, True))

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta de Bancos en el debe : Primer Capital
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe, pnImporte
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod


'Guardamos la Cuenta de la Pendiente que se elimina
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnImporte * -1
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod

'Pasamos a activo el estado de la cuenta aperturada
'oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIfCod, , , Format(ldFecha, gsFormatoFecha), Format(ldFecha, gsFormatoFecha), , , gEstadoCtaIFActiva
oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIFCod, , , Format(pdFecApe, gsFormatoFecha), Format(pdFecApe, gsFormatoFecha), , , gEstadoCtaIFActiva
oDMov.InsertaMovRef lnMovNro, pnMovNroRef

If pnTpoInversion = 3 Then 'Si es Fonfo Mutuo
                                           
        oDMov.InsertaInversion lnMovNro, psOpecod, psPersCod, psIFTpo, psCtaIFCod, psCtaIFDesc, _
                                   pnTpoInversion, pnMoneda, pnCondicion, pnImporte, 0, _
                                   pdFecApe, pdFecApe, pdFecApe, 0, 0, _
                                   psPerCodTrans, psIFTipoTrans, psCtaIfCodTrans, _
                                   psMovDesc, psMovNro
        
        oDMov.actualizaFondoMutuoMov lnMovNro, pnMovNroRef, psIFTpo, psPersCod, psCtaIFCod
Else
        oDMov.InsertaInversion lnMovNro, psOpecod, psPersCod, psIFTpo, psCtaIFCod, psCtaIFDesc, _
                                   pnTpoInversion, pnMoneda, pnCondicion, pnImporte, pnPlazo, _
                                   pdFecApe, pdFecVenc, pdFecApe, pnTasa, pnIntPactado, _
                                   psPerCodTrans, psIFTipoTrans, psCtaIfCodTrans, _
                                   psMovDesc, psMovNro
End If
If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"
'PASI20150921 ERS0712015
If pnTpoInversion = 6 Then
    oDMov.InsertaMov psMovNroACAdd, psOpecod, psMovDesc
    lnMovNroACAdd = oDMov.GetnMovNro(psMovNroACAdd)
    oDMov.InsertaMovCont lnMovNroACAdd, pnImporte, "0", "0"
    lnMovItem = 0
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNroACAdd, lnMovItem, "83" & IIf(Mid(psOpecod, 3, 1) = gMonedaExtranjera, "2", "1") & "1", pnImporte
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNroACAdd, lnMovItem, "84" & IIf(Mid(psOpecod, 3, 1) = gMonedaExtranjera, "2", "1") & "90701", pnImporte * -1
    If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
        oDMov.GeneraMovME lnMovNroACAdd, psMovNroACAdd
    End If
    oDMov.ActualizaSaldoMovimiento psMovNroACAdd, "+"
    oDMov.InsertaMovRef lnMovNroACAdd, lnMovNro
End If
'END PASI

oDMov.CommitTrans
lbTrans = False

GrabaConfInversion = 0
Set oDMov = Nothing

Exit Function
MsgError:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaConfInversion", lsMsgErr & " Error en Grabación de Confirmación de Inversiones "
End Function
Public Function obtenerInversionesVigentes(psOpecod As String, pdFecha As Date, psTipoInversion As String) As Recordset
    On Error GoTo MsgError
    
    Dim rs As New ADODB.Recordset
    
    sql = "exec stp_sel_InversionVigentes '" & psOpecod & "','" & Format(pdFecha, "yyyymmdd") & "','" & psTipoInversion & "'"
    
    Set rs = oConect.CargaRecordSet(sql)
    Set obtenerInversionesVigentes = rs
    
    
    Exit Function
MsgError:
    lsMsgErr = Err.Description
    Err.Raise vbObjectError + 100, "obtenerInversionesConf", lsMsgErr & " Error Obtener Inversiones Vigentes"
End Function
Public Function GrabaCancelInversion(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
                                    ByVal psCtaDebe As String, ByVal pnImporte As Currency, _
                                    ByVal psObjetoCodDebe As String, _
                                    ByVal psCtaHaberInt As String, ByVal pnImporteInt As Currency, _
                                    ByVal psCtaHaberIntCal As String, ByVal pnImporteIntCalc As Currency, _
                                    ByVal psCtaHaberCapital As String, ByVal pnImporteCapital As Currency, _
                                    ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psCtaIFCod As String, _
                                    ByVal pdCtaIFCap As Date, ByVal pnMovNroRef As Long, _
                                    ByVal pnTpoInversion As Integer, ByVal pnMoneda As Integer, ByVal pnCondicion As Integer, _
                                    ByVal pnPlazo As Integer, ByVal pdFecApe As Date, ByVal pdFecVenc As Date, ByVal pnTasa As Double, _
                                    ByVal pnIntPactado As Currency, ByVal psCtaIFDesc As String, ByVal pnDias As Integer, ByVal pnCapital As Currency, _
                                    ByVal pnValorCuota As Double, ByVal pnNroCuotas As Currency, ByVal pnImpRenta As Currency, Optional psCtaAntiImpRenta As String = "", Optional ByVal psMovNroACAdd As String = "") As Integer
                                    'agrego psMovNroACAdd PASI20150921 ERS0712015
            
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lnMovNroACAdd As Long 'PASI20150921 ERS0712015
Set oDMov = New DMov
On Error GoTo MsgError



GrabaCancelInversion = 1

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpecod, psMovDesc
lnMovNro = oDMov.GetnMovNro(psMovNro)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, pnImporte
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psObjetoCodDebe, 4, 13), Mid(psObjetoCodDebe, 1, 2), Mid(psObjetoCodDebe, 18, 10)
            
'Impuesto a la renta
If pnImpRenta > 0 And pnTpoInversion = 3 Then
    
'    Dim clsTC As COMDConstSistema.NCOMTipoCambio
'    Dim nTC As Double
'    Set clsTC = New COMDConstSistema.NCOMTipoCambio
'    If pnMoneda = 2 Then
'        nTC = clsTC.EmiteTipoCambio(pdFecVenc, TCFijoDia)
'    Else
'        nTC = 1
'    End If

    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, "45" & CStr(pnMoneda) & "40902", pnImpRenta
    
    'Contraparte para cuadrar el asiento; Cuenta Haber solo cuando hay perdida
     If Left(psCtaHaberIntCal, 2) = "43" Then
        lnMovItem = lnMovItem + 1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaAntiImpRenta, -pnImpRenta
      End If
End If
            
'guardamos la cuenta en el Haber
If pnImporteInt <> 0 And pnTpoInversion <> 3 Then
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaberInt, pnImporteInt * -1
End If


If pnImporteIntCalc <> 0 Then
    lnMovItem = lnMovItem + 1:
    If Left(psCtaHaberIntCal, 2) <> "43" Then
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaberIntCal, (pnImporteIntCalc + pnImpRenta) * -1
    Else
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaberIntCal, pnImporteIntCalc
    End If
End If


If pnImporteCapital > 0 And psCtaHaberCapital <> "" Then
    lnMovItem = lnMovItem + 1:
    lnMovOrden = 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaberCapital, pnImporteCapital * -1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod
    
End If

oDMov.InsertaMovRef lnMovNro, pnMovNroRef
If pnTpoInversion <> 3 Then
    oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIFCod, , , Format(pdCtaIFCap, gsFormatoFecha), , pnPlazo, Format(pdCtaIFCap, gsFormatoFecha), gEstadoCtaIFCancelada, (pnImporteInt + pnImporteIntCalc), psMovNro
    oDMov.InsertaInversion lnMovNro, psOpecod, psPersCod, psIFTpo, psCtaIFCod, psCtaIFDesc, _
                                       pnTpoInversion, pnMoneda, pnCondicion, pnCapital, pnDias, _
                                       pdFecApe, pdFecVenc, pdFecApe, pnTasa, pnIntPactado, _
                                       Mid(psObjetoCodDebe, 4, 13), Mid(psObjetoCodDebe, 1, 2), Mid(psObjetoCodDebe, 18, 10), _
                                       psMovDesc, psMovNro, (pnImporteIntCalc + pnImporteInt)
 Else
    oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIFCod, , , Format(pdCtaIFCap, gsFormatoFecha), , 0, Format(pdCtaIFCap, gsFormatoFecha), gEstadoCtaIFCancelada, 0, psMovNro
    oDMov.InsertaInversion lnMovNro, psOpecod, psPersCod, psIFTpo, psCtaIFCod, psCtaIFDesc, _
                                       pnTpoInversion, pnMoneda, pnCondicion, pnCapital, 0, _
                                       pdFecApe, pdFecVenc, pdFecApe, 0, 0, _
                                       Mid(psObjetoCodDebe, 4, 13), Mid(psObjetoCodDebe, 1, 2), Mid(psObjetoCodDebe, 18, 10), _
                                       psMovDesc, psMovNro, 0

    oDMov.InsertaInversionFondosMutuos lnMovNro, pdFecVenc, psPersCod, Format(psIFTpo, "00"), psCtaIFCod, pnImporteCapital, pnValorCuota, pnNroCuotas, pnImpRenta
 End If
If Mid(psOpecod, 3, 1) = "2" Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If

oDMov.ActualizaSaldoMovimiento psMovNro, "+"
'PASI20150921 ERS0712015
If pnTpoInversion = 6 Then
    oDMov.InsertaMov psMovNroACAdd, psOpecod, psMovDesc
    lnMovNroACAdd = oDMov.GetnMovNro(psMovNroACAdd)
    oDMov.InsertaMovCont lnMovNroACAdd, pnImporte, "0", "0"
    lnMovItem = 0
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNroACAdd, lnMovItem, "84" & IIf(Mid(psOpecod, 3, 1) = gMonedaExtranjera, "2", "1") & "90701", pnImporteCapital
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNroACAdd, lnMovItem, "83" & IIf(Mid(psOpecod, 3, 1) = gMonedaExtranjera, "2", "1") & "1", pnImporteCapital * -1
    If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
        oDMov.GeneraMovME lnMovNroACAdd, psMovNroACAdd
    End If
    oDMov.InsertaMovRef lnMovNroACAdd, lnMovNro
    oDMov.ActualizaSaldoMovimiento psMovNroACAdd, "+"
End If
'END PASI
oDMov.CommitTrans
lbTrans = False

GrabaCancelInversion = 0
Set oDMov = Nothing
Exit Function
MsgError:
lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaCancelInversion", lsMsgErr & " Error Graba Cancelacion de Inversion"
End Function
Public Function GrabaProvInversion(ByVal psMovNro As String, ByVal psOpecod As String, ByVal psMovDesc As String, _
                                   ByVal psCtaDebe As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                   ByVal pnInteres As Currency, ByVal psPersCod As String, ByVal psIFTpo As String, _
                                   ByVal psCtaIFCod As String, ByVal pdCtaIFInt As Date, ByVal pnMovNroRef As Long, _
                                   ByVal pbFinMes As Boolean, ByVal pnTpoInversion As Integer, _
                                   ByVal pnCapFndMutuo As Currency, ByVal pnValorCuota As Double, ByVal pnNroCuotas As Currency) As Integer
                         

Dim lnMovItem As Long
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lnInteres As Currency
Dim oCtaIf As NCajaCtaIF
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim bFinMes As Boolean
Dim dFecFinMes As Date
Set oDMov = New DMov
On Error GoTo MsgError

GrabaProvInversion = 1

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpecod, psMovDesc
lnMovNro = oDMov.GetnMovNro(psMovNro)

If pbFinMes = True Or pnTpoInversion <> 3 Then
    oDMov.InsertaMovCont lnMovNro, pnInteres, "0", "0"
    'guardamos la cuenta en el debe
    lnMovItem = 0: lnMovOrden = 0
    lnMovItem = lnMovItem + 1:
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, pnInteres
    lnMovOrden = lnMovOrden + 1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod
    'guardamos la cuenta en el Haber
    
    If Left(psCtaDebe, 2) <> "43" Then
        lnMovItem = lnMovItem + 1: lnMovOrden = 0
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnInteres * -1
        lnMovOrden = lnMovOrden + 1
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod
    End If
End If

If pnTpoInversion <> 3 Then
    oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIFCod, , , , Format(pdCtaIFInt, gsFormatoFecha), , , , pnImporte, psMovNro
    oDMov.actualizaInteresInversion pnMovNroRef, "42" + Mid(psOpecod, 3, 1) + "302", pnImporte
Else
    oDMov.actualizaInteresInversion pnMovNroRef, "42" + Mid(psOpecod, 3, 1) + "302", 0
    oDMov.InsertaInversionFondosMutuos lnMovNro, pdCtaIFInt, psPersCod, Format(psIFTpo, "00"), psCtaIFCod, pnCapFndMutuo, pnValorCuota, pnNroCuotas, 0
End If

If pbFinMes = True Or pnTpoInversion <> 3 Then
    If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
        oDMov.GeneraMovME lnMovNro, psMovNro
    End If
End If
oDMov.CommitTrans
lbTrans = False

GrabaProvInversion = 0
Set oDMov = Nothing
Set oCtaIf = Nothing
Exit Function
MsgError:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaProvInversion", Err.Description & " Error Grabación de Provision de Inversiones"
End Function
Public Function obtenerInversionesListaExtornar(psOpecod As String, psOpeCodVig As String, pdFechaIni As Date, pdFechaFin As Date, psTipoInversion As String) As Recordset
    On Error GoTo MsgError
    
    Dim rs As New ADODB.Recordset
    
    sql = "exec stp_sel_InversionListaExtorna '" & psOpecod & "','" & psOpeCodVig & "','" & Format(pdFechaIni, "yyyymmdd") & "','" & Format(pdFechaFin, "yyyymmdd") & "','" & psTipoInversion & "'"
    
    Set rs = oConect.CargaRecordSet(sql)
    Set obtenerInversionesListaExtornar = rs
        
    Exit Function
MsgError:
    lsMsgErr = Err.Description
    Err.Raise vbObjectError + 100, "obtenerInversionesListaExtornar", lsMsgErr & " Error Obtener Inversiones Listado a Extornar"
End Function
Public Function GrabaExtornoInversiones(ByVal psMovNro As String, ByVal pnMovNroExt As Long, ByVal psMovNroExt As String, _
                                        ByVal psOpecod As String, ByVal psMovDesc As String, ByVal pnCalculado As Currency, _
                                        ByVal lbEliminaMov As Boolean, ByVal psObjetoCod As String, ByVal pnInteres As Currency, _
                                        ByVal psOpeCodExt As String, Optional ByVal pnTpoInversion As Integer = 0)

On Error GoTo MsgError
Dim oMov As DMov
Dim lbTrans As Boolean
Dim oChq As New NDocRec

Set oMov = New DMov
oMov.BeginTrans
lbTrans = True
    
    oMov.ExtornaMovimiento psMovNro, pnMovNroExt, psOpecod, psMovDesc, lbEliminaMov, psMovNro
    Select Case psOpecod
        Case "421306", "422306" 'Apertura
            oMov.ActualizaCtaIF Mid(psObjetoCod, 4, 13), Left(psObjetoCod, 2), Mid(psObjetoCod, 18), , , , , , , gEstadoCtaIFAnulada
            oMov.actualizaEstadoInversion pnMovNroExt, psOpeCodExt, 0 'Inactiva el estado
            If pnTpoInversion = 3 Then
                oMov.actualizaEstadoFondoMutuo pnMovNroExt, Left(psObjetoCod, 2), Mid(psObjetoCod, 4, 13), Mid(psObjetoCod, 18)
            End If
        Case "421307", "422307" 'Confirmacion
            oMov.ActualizaCtaIF Mid(psObjetoCod, 4, 13), Left(psObjetoCod, 2), Mid(psObjetoCod, 18), , , , , , , gEstadoCtaIFRegistrada
            oMov.actualizaEstadoInversion pnMovNroExt, psOpeCodExt, 0 'Inactiva el estado
        Case "421308", "422308" 'Cancelacion
            oMov.ActualizaCtaIF Mid(psObjetoCod, 4, 13), Left(psObjetoCod, 2), Mid(psObjetoCod, 18), , , , , , , gEstadoCtaIFActiva, pnInteres
            oMov.actualizaEstadoInversion pnMovNroExt, psOpeCodExt, 0 'Inactiva el estado
            If pnTpoInversion = 3 Then
                oMov.actualizaEstadoFondoMutuo pnMovNroExt, Left(psObjetoCod, 2), Mid(psObjetoCod, 4, 13), Mid(psObjetoCod, 18)
            End If
        Case "421309", "422309" 'Provision
             Dim rs As New Recordset
             Set rs = oMov.getUltFecProvInversion(pnMovNroExt, "42" + Mid(psOpecod, 3, 1) + "304", Left(psObjetoCod, 2), Mid(psObjetoCod, 4, 13), Mid(psObjetoCod, 18))
             
             oMov.ActualizaCtaIF Mid(psObjetoCod, 4, 13), Left(psObjetoCod, 2), Mid(psObjetoCod, 18), , , , Left(rs!cMovNro, 8), , , gEstadoCtaIFActiva, pnInteres - pnCalculado, rs!cMovNro
             oMov.actualizaInteresInversion pnMovNroExt, "42" + Mid(psOpecod, 3, 1) + "302", pnInteres - pnCalculado
             If pnTpoInversion = 3 Then
                oMov.actualizaEstadoFondoMutuo pnMovNroExt, Left(psObjetoCod, 2), Mid(psObjetoCod, 4, 13), Mid(psObjetoCod, 18)
             End If
    End Select
   
    
oMov.CommitTrans
lbTrans = False
Set oMov = Nothing
Set oChq = Nothing
Exit Function
MsgError:
    lsMsgErr = Err.Description
    If lbTrans Then oMov.RollbackTrans
    Err.Raise Err.Number, "GrabaExtornoInversiones", lsMsgErr
End Function
Public Function getInversionesListadoMantenimiento(psMoneda As String, pdFechaIni As Date, pdFechaFin As Date, psTipoInversion As String) As Recordset
    On Error GoTo MsgError
    
    Dim rs As New ADODB.Recordset
    
    sql = "exec stp_sel_InversionesListadoMant '" & psMoneda & "','" & Format(pdFechaIni, "yyyymmdd") & "','" & Format(pdFechaFin, "yyyymmdd") & "','" & psTipoInversion & "'"
    
    Set rs = oConect.CargaRecordSet(sql)
    Set getInversionesListadoMantenimiento = rs
        
    Exit Function
MsgError:
    lsMsgErr = Err.Description
    Err.Raise vbObjectError + 100, "getInversionesListadoMantenimiento", lsMsgErr & " Error Obtener Inversiones Listado Mantenimiento"
End Function
Public Function getInversionDatos(pnMovNro As Long, psOpecod As String) As Recordset
    On Error GoTo MsgError
    
    Dim rs As New ADODB.Recordset
    
    sql = "exec stp_sel_InversionDatos " & pnMovNro & ",'" & psOpecod & "'"
    
    Set rs = oConect.CargaRecordSet(sql)
    Set getInversionDatos = rs
        
    Exit Function
MsgError:
    lsMsgErr = Err.Description
    Err.Raise vbObjectError + 100, "getInversionDatos", lsMsgErr & " Error Obtener Datos de la Inversion"
End Function
Public Sub actualizaInversionDatos(ByVal pnMovNro As Long, ByVal psOpecod As String, ByVal psPerCod As String, _
                                    ByVal psIFTipo As String, ByVal psCtaIFCod As String, ByVal psCtaIFDesc As String, _
                                    ByVal pnPlazo As Integer, ByVal pdFecApe As Date, ByVal pdFecVenc As Date, _
                                    ByVal pnImporte As Currency, ByVal pnTasa As Double, ByVal pnIntPactado As Currency, _
                                    ByVal psMovDesc As String, ByVal psMovNro As String, ByVal pnTpoInversion As Integer, _
                                    ByVal pnCondicion As Integer, Optional pnValorCuota As Double = 0#, Optional pnNroCuotas As Currency = 0#)
    Dim oMov As DMov
    Dim lbTrans As Boolean
    

    Set oMov = New DMov
    oMov.BeginTrans
    lbTrans = True
    
            oMov.ActualizaCtaIF psPerCod, psIFTipo, psCtaIFCod, psCtaIFDesc, Format(pdFecApe, "mm/dd/yyyy"), , , pnPlazo, Format(pdFecVenc, "mm/dd/yyyy"), IIf(Right(psOpecod, 1) = "1", gEstadoCtaIFRegistrada, gEstadoCtaIFActiva)
            oMov.ActualizaCuentaIFInteres psPerCod, psIFTipo, psCtaIFCod, pdFecApe, pnPlazo, pnTasa, psMovNro
            oMov.actualizaInversionDatos pnMovNro, psOpecod, psCtaIFDesc, pnTpoInversion, pnCondicion, pnImporte, pnPlazo, _
                                         pdFecApe, pdFecVenc, pnTasa, pnIntPactado, psMovDesc
            If pnTpoInversion = 3 Then
                oMov.actualizaInversionFondosMutuos pnMovNro, pdFecApe, psPerCod, psIFTipo, psCtaIFCod, pnImporte, pnValorCuota, pnNroCuotas, 0
            End If
   
    
oMov.CommitTrans
lbTrans = False
Set oMov = Nothing

Exit Sub
MsgError:
    lsMsgErr = Err.Description
    If lbTrans Then oMov.RollbackTrans
    Err.Raise Err.Number, "actualizaInversionDatos", lsMsgErr & " Error Actualizar Datos de Inversion"
End Sub
'JACA END*************************************************************
'***Agregado por ELRO el 20121001, según OYP-RFC106-2012
Public Function registrarOperacionesTranferenciaNegocio(ByVal psMovNro As String, _
                                                        ByVal psOpecod As String, _
                                                        ByVal psMovDesc As String, _
                                                        ByVal pnMovNroOpe As Long, _
                                                        ByVal pnImporte As Currency, _
                                                        Optional ByVal pnTCBanco As Currency = 0) As Long
Dim oDMov As New DMov
Dim lbTrans As Boolean
Dim rsMovOpe As New ADODB.Recordset
Dim rsPlaNeg As New ADODB.Recordset

Dim lsMonedaOpe, lsAgeCodOpe, lsMovNroOpe, lsCtaContCod  As String

Dim lsBC As String
Dim lnMovNro As Long

Dim lnMovItem, lnMovOrden As Integer

oDMov.BeginTrans
lbTrans = True

oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
    
Set rsMovOpe = oDMov.devolverMovimientoTransferencia(pnMovNroOpe)

If Not (rsMovOpe.BOF And rsMovOpe.EOF) Then
    lsAgeCodOpe = Mid(rsMovOpe!cMovNro, 18, 2)
    lsMonedaOpe = Mid(rsMovOpe!cCtaCod, 9, 1)
    lsBC = rsMovOpe!cCtaIFSubCta
    lnMovItem = 0: lnMovOrden = 0:
    Do While Not rsMovOpe.EOF
     
     Set rsPlaNeg = oDMov.DevolverPlantillaNegocioTransferencia(rsMovOpe!cOpeCod, rsMovOpe!nConceptoCod, rsMovOpe!nPersoneria)
     If Not (rsPlaNeg.BOF And rsPlaNeg.EOF) Then

        Do While Not rsPlaNeg.EOF
            lsCtaContCod = rsPlaNeg!cCtaContCod
            'EJVG20130920 *** x BCRP
            If lsCtaContCod = "11M301BC" And rsMovOpe!cPersCodIF = "1090100822183" Then
                lsCtaContCod = "11M20101"
            End If
            'END EJVG *******
            If InStr(lsCtaContCod, "M") > 0 Then
                lsCtaContCod = Replace(lsCtaContCod, "M", lsMonedaOpe)
            End If
            
            If InStr(lsCtaContCod, "AG") > 0 Then
                lsCtaContCod = Replace(lsCtaContCod, "AG", lsAgeCodOpe)
            End If
            
            If InStr(lsCtaContCod, "BC") > 0 Then
                lsCtaContCod = Replace(lsCtaContCod, "BC", lsBC)
            End If
            
            If InStr(lsCtaContCod, "CJ") > 0 Then
                lsCtaContCod = Replace(lsCtaContCod, "CJ", oDMov.devolverVariableInstitucionFinanciera(rsMovOpe!cCtaCod, gTpoIFCmac))
            End If
            
            If InStr(lsCtaContCod, "CR") > 0 Then
                lsCtaContCod = Replace(lsCtaContCod, "CR", oDMov.devolverVariableInstitucionFinanciera(rsMovOpe!cCtaCod, gTpoIFCrac))
            End If
            
            If InStr(lsCtaContCod, "CO") > 0 Then
                lsCtaContCod = Replace(lsCtaContCod, "CO", oDMov.devolverVariableInstitucionFinanciera(rsMovOpe!cCtaCod, gTpoIFCooperativa))
            End If
            
            If InStr(lsCtaContCod, "ED") > 0 Then
                lsCtaContCod = Replace(lsCtaContCod, "ED", oDMov.devolverVariableInstitucionFinanciera(rsMovOpe!cCtaCod, gTpoIFEDPYME))
            End If

           lnMovItem = lnMovItem + 1
           lnMovOrden = lnMovOrden + 1
            If rsPlaNeg!cOpeCtaDH = "D" Then 'guardamos la cuenta en el debe
                oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCtaContCod, rsMovOpe!nMonto
                oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
                oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, rsMovOpe!cPersCodIF, rsMovOpe!cIFTpo, rsMovOpe!cCtaIFCod
            Else 'guardamos la cuenta en el haber
                oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCtaContCod, rsMovOpe!nMonto * -1
            End If
            rsPlaNeg.MoveNext
        Loop
     Else
        MsgBox "No existe plantilla del negocio para esta operación, según los criterios " & Chr(13) & " Código Operación: " & rsMovOpe!cOpeCod & ", Concepto Operación: " & rsMovOpe!nConceptoCod & ", Personería Operación: " & rsMovOpe!nPersoneria & Chr(13) & "Comuniquese con Contabilidad."
        If lbTrans Then oDMov.RollbackTrans
     Exit Function
     End If
        
    rsMovOpe.MoveNext
    Loop
End If
oDMov.InsertaMovRef lnMovNro, pnMovNroOpe


If Mid(psOpecod, 3, 1) = 2 Then
    oDMov.GeneraMovME lnMovNro, psMovNro, pnTCBanco
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"

oDMov.CommitTrans
registrarOperacionesTranferenciaNegocio = lnMovNro
lbTrans = False
Set oDMov = Nothing

Exit Function
MsgError:
    lsMsgErr = Err.Description
    If lbTrans Then oDMov.RollbackTrans
    Err.Raise Err.Number, "registrarOperacionesTranferenciaNegocio", lsMsgErr & " Error Actualizar Datos de Inversion"

End Function
'***Fin Agregado por ELRO el 20121001*******************
'EJVG20121019 ***
Public Sub GrabarFlujoCajaProyectado(ByVal pnAnio As Integer, pnMoneda As Moneda, ByVal psUser As String, pdFecha As Date, ByVal rsFCP As ADODB.Recordset)
    Dim oFlujoCaja As New DCajaGeneral
    Dim lnNuevoCod As Integer
    Dim lbTrans As Boolean
    lbTrans = False
    
    On Error GoTo MsgError
    
    lnNuevoCod = oFlujoCaja.obtenerNuevoCodFCP(pnAnio, pnMoneda)
    oFlujoCaja.dBeginTrans
    
    If lnNuevoCod > 1 Then
        Call oFlujoCaja.ActualizaFlujoCajaProyectado(pnAnio, pnMoneda, lnNuevoCod)
    Else
        Call oFlujoCaja.InsertaFlujoCajaProyectado(pnAnio, pnMoneda, lnNuevoCod, psUser, pdFecha)
    End If

    Do While Not rsFCP.EOF
        Call oFlujoCaja.InsertaFlujoCajaProyectadoDet(pnAnio, pnMoneda, lnNuevoCod, 1, rsFCP![FlujoId], rsFCP![Enero])
        Call oFlujoCaja.InsertaFlujoCajaProyectadoDet(pnAnio, pnMoneda, lnNuevoCod, 2, rsFCP![FlujoId], rsFCP![Febrero])
        Call oFlujoCaja.InsertaFlujoCajaProyectadoDet(pnAnio, pnMoneda, lnNuevoCod, 3, rsFCP![FlujoId], rsFCP![Marzo])
        Call oFlujoCaja.InsertaFlujoCajaProyectadoDet(pnAnio, pnMoneda, lnNuevoCod, 4, rsFCP![FlujoId], rsFCP![Abril])
        Call oFlujoCaja.InsertaFlujoCajaProyectadoDet(pnAnio, pnMoneda, lnNuevoCod, 5, rsFCP![FlujoId], rsFCP![Mayo])
        Call oFlujoCaja.InsertaFlujoCajaProyectadoDet(pnAnio, pnMoneda, lnNuevoCod, 6, rsFCP![FlujoId], rsFCP![Junio])
        Call oFlujoCaja.InsertaFlujoCajaProyectadoDet(pnAnio, pnMoneda, lnNuevoCod, 7, rsFCP![FlujoId], rsFCP![Julio])
        Call oFlujoCaja.InsertaFlujoCajaProyectadoDet(pnAnio, pnMoneda, lnNuevoCod, 8, rsFCP![FlujoId], rsFCP![Agosto])
        Call oFlujoCaja.InsertaFlujoCajaProyectadoDet(pnAnio, pnMoneda, lnNuevoCod, 9, rsFCP![FlujoId], rsFCP![Septiembre])
        Call oFlujoCaja.InsertaFlujoCajaProyectadoDet(pnAnio, pnMoneda, lnNuevoCod, 10, rsFCP![FlujoId], rsFCP![Octubre])
        Call oFlujoCaja.InsertaFlujoCajaProyectadoDet(pnAnio, pnMoneda, lnNuevoCod, 11, rsFCP![FlujoId], rsFCP![Noviembre])
        Call oFlujoCaja.InsertaFlujoCajaProyectadoDet(pnAnio, pnMoneda, lnNuevoCod, 12, rsFCP![FlujoId], rsFCP![Diciembre])
        rsFCP.MoveNext
    Loop
    
    oFlujoCaja.dCommitTrans
    lbTrans = True
    
    Set oFlujoCaja = Nothing
    Exit Sub
MsgError:
    If Not lbTrans Then
        oFlujoCaja.dRollbackTrans
    End If
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Error"
End Sub
'END EJVG *******
'EJVG20131125 ***
Public Function GrabaPagoProveedor_NEW(psMovNro As String, psOpecod As String, psMovDesc As String, psCtaContDebeB As String, psCtaContDebeS As String, _
                                   psCtaContHaber As String, pnImporteB As Currency, pnImporteS As Currency, psPersCod As String, psTpoIF As String, psPersCodIf As String, psCtaIf As String, _
                                   prsBilletaje As ADODB.Recordset, psDocTpo As String, psDocNro As String, psFechaDoc As String, psDocVoucher As String, rsDocs As ADODB.Recordset, psCuentaAho As String, _
                                   ByVal psCtaDiferencia As String, ByVal pnMontoDiferencia As Currency, Optional ByVal gbBitCentral As Boolean, Optional pbGrabaOpeNegocio As Boolean = False, Optional pnRetAct As Currency = 0, Optional lsCtaITFD As String, Optional lsCtaITFH As String, Optional nTasaITF As Double = 0, Optional bIncluyeNC As Boolean = False, _
                                   Optional bAfectoITF As Boolean = True, Optional lnITFDescCuentas As Double, Optional pnImporteRH As Currency, Optional psCtaContDebeRH As String, Optional psCtacontDebeRHJ As String, Optional pnImporteRHJ As Currency, Optional psCtaContDebeSegu As String, Optional pnImporteSegu As Currency, _
                                   Optional psCtaContDebePagoVarios As String, Optional pnImportePagoVarios As Currency, Optional nTipoPago As Integer = 0, Optional ByVal psCtaSaf As String = "", Optional ByVal lsCtaCodCore As String = "", Optional ByVal psCtaContLeasing As String = "", Optional ByVal pnItmActual As Integer, Optional ByVal pnMovNroProv As Long = 0, _
                                   Optional ByVal psCtaContPenalidadH As String = "") As Integer
Dim lnMovItem  As Long
Dim lnMovOrden As Integer
Dim lnMovNro As Long
Dim lbTrans  As Boolean
Dim lsMovNro As String
Dim lsCadBol As String
Dim oDMov    As DMov
Set oDMov = New DMov

Dim oConst As NConstSistemas
Set oConst = New NConstSistemas
Dim oImp As DImpuesto
Set oImp = New DImpuesto
Dim oFun As NContFunciones
Set oFun = New NContFunciones
Dim lbBitReten As Boolean
Dim lsCtaReten As String
Dim lsCtaRetenConvesion As String
Dim lsCtaRetenCambio As String
Dim lbBCAR As Boolean
Dim lnTasaImp As Currency
Dim lnIngresos As Currency
Dim lnRetencion As Currency
Dim lnTopeRetencion As Currency
Dim lnRetAct As Currency
Dim lnRetActME As Currency
Dim lsRetDocNro As String
Dim lMN As Boolean
Dim lsMovNroNegocio As String
Dim lbGraboNegocio As Boolean
Dim rsPenalidad As ADODB.Recordset
Dim lnPenalidad As Currency

On Error GoTo GrabaPagoProveedorErr

GetTipCambio GetFechaMov(psMovNro, True)

If Mid(psCtaContHaber, 3, 1) = Moneda.gMonedaExtranjera Then
    lMN = False
Else
    lMN = True
End If

lbBitReten = IIf(oConst.LeeConstSistema(gConstSistBitRetencion6Porcent) = 1, True, False)
If lbBitReten And pnRetAct <> 0 Then
   lsRetDocNro = oFun.GeneraDocNro(TpoDocAlmacenComprdeRetencion6)
    lsCtaReten = oConst.LeeConstSistema(gConstSistCtaRetencion6Porcent)
End If

oDMov.BeginTrans
lbTrans = True
GrabaPagoProveedor_NEW = 0

oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu, "0", "0"

'ALPA 20110909*****************************************************
If Len(Trim(psCtaSaf)) > 0 Then
Dim oCtSaldo As DCtaSaldo
Set oCtSaldo = New DCtaSaldo
Call oCtSaldo.InsertarDesembolso_Log(psCtaSaf, lsCtaCodCore, lnMovNro)
End If
'******************************************************************

'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
If pnImporteB <> 0 Then
   lnMovItem = lnMovItem + 1
   'ALPA 20090317******************************************************
   'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeB, pnImporteB
   'ALPA 20110912
   'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeB, pnImporteB, nTipoPago
   If Len(Trim(psCtaSaf)) = 0 Then
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeB, pnImporteB, nTipoPago
   Else
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContLeasing, pnImporteB, nTipoPago
   End If
   '*******************************************************************
End If
If pnImporteS <> 0 Then
   lnMovItem = lnMovItem + 1
   'ALPA 20090317******************************************************
   'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeS, pnImporteS
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeS, pnImporteS, nTipoPago
   '*******************************************************************
End If
'JEOM
If pnImporteRH <> 0 Then
   lnMovItem = lnMovItem + 1
   'ALPA 20090317******************************************************
   'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeRH, pnImporteRH
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeRH, pnImporteRH, nTipoPago
   '*******************************************************************
End If
If pnImporteRHJ <> 0 Then
   lnMovItem = lnMovItem + 1
   'ALPA 20090317******************************************************
   'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtacontDebeRHJ, pnImporteRHJ
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtacontDebeRHJ, pnImporteRHJ, nTipoPago
   '*******************************************************************
End If
If pnImporteSegu <> 0 Then
   lnMovItem = lnMovItem + 1
   'ALPA 20090317******************************************************
   'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeSegu, pnImporteSegu
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeSegu, pnImporteSegu, nTipoPago
   '*******************************************************************
End If

If pnImportePagoVarios <> 0 Then
   lnMovItem = lnMovItem + 1
   'ALPA 20090317*******************************************************
   'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebePagoVarios, pnImportePagoVarios
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebePagoVarios, pnImportePagoVarios, nTipoPago
   '********************************************************************
End If

'Agregado

lnRetAct = 0
lnRetActME = 0
If lbBitReten Then
   lnRetAct = pnRetAct
   ''   lnRetActME = Round(lnRetAct / gnTipCambioPonderado, 2)
   lnRetActME = Round(lnRetAct / gnTipCambioPonderado, 2)
   lnMovItem = lnMovItem + 1
   'ALPA 20090317*********************************************************
   'oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCtaReten, lnRetAct * -1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCtaReten, lnRetAct * -1, nTipoPago
   oDMov.InsertaMovDoc lnMovNro, TpoDocAlmacenComprdeRetencion6, lsRetDocNro, psFechaDoc
   '***********************************************************************
End If

'guardamos el billetaje si este posee
If Not prsBilletaje Is Nothing Then
   oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, prsBilletaje, psCtaContHaber, "H"
Else
    'guardamos la cuenta del haber
'    If (pnImporteB + pnImporteS - lnRetAct) <> 0 Then
    If (pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu + pnImportePagoVarios - IIf(lMN, lnRetAct, lnRetActME)) <> 0 Then
        lnMovItem = lnMovItem + 1: lnMovOrden = 0
'        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber, (pnImporteB + pnImporteS - IIf(lnRetAct > 0, lnRetAct, 0)) * -1
        'ALPA 20090317*******************************************************
        'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber, (pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu + pnImportePagoVarios - IIf(lnRetAct + lnRetActME > 0, IIf(lMN, lnRetAct, lnRetActME), 0)) * -1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber, (pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu + pnImportePagoVarios - IIf(lnRetAct + lnRetActME > 0, IIf(lMN, lnRetAct, lnRetActME), 0)) * -1, nTipoPago
        '********************************************************************
    End If
    If psPersCodIf <> "" Then
        'guardamos la Cta de la IF si es que utiliza esta
        lnMovOrden = lnMovOrden + 1
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCodIf, psTpoIF, psCtaIf
        
    End If
End If
If psDocNro <> "" Then
   oDMov.InsertaMovDoc lnMovNro, psDocTpo, psDocNro, psFechaDoc
End If
If psDocVoucher <> "" Then
   oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocVoucher, GetFechaMov(psMovNro, False)
End If

oDMov.InsertaMovGasto lnMovNro, psPersCod, "0"

If pnMontoDiferencia <> 0 Then
    lnMovItem = lnMovItem + 1
    'oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDiferencia, pnMontoDiferencia * -1
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDiferencia, pnMontoDiferencia * -1, nTipoPago
End If

'Referenciamos todos los Documentos que se Pagan
If Not rsDocs Is Nothing Then '*** PEAC 20110218
If bIncluyeNC Or rsDocs!nDocTpo = 7 Then
    '   OJO: Pago Total
'    If Not rsDocs Is Nothing Then
       Do While Not rsDocs.EOF
          If rsDocs!Itm = "1" And rsDocs.Fields(0) = pnItmActual Then
            'ALPA 20090317***********************************************
             'oDMov.InsertaMovRef lnMovNro, rsDocs!nMovNro
             oDMov.InsertaMovRef lnMovNro, rsDocs!nMovNro, , nTipoPago
             '***********************************************************
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeB, 0, False
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeS, 0, False
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeRH, 0, False
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtacontDebeRHJ, 0, False
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeSegu, 0, False
             oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebePagoVarios, 0, False
         
          End If
          rsDocs.MoveNext
       Loop
    'End If
Else
    ' Puede ser pago parcial
    
    Dim nMontoPago As Currency
    'If Not rsDocs Is Nothing Then
        nMontoPago = pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu + pnImportePagoVarios
       
       'Do While Not rsDocs.EOF And nMontoPago > 0 '*** PEAC 20110301
       Do While Not rsDocs.EOF
       
          If rsDocs!Itm = "1" And rsDocs.Fields(0) = pnItmActual Then
             'ALPA 20090317************************************************
             'oDMov.InsertaMovRef lnMovNro, rsDocs!nMovNro
             oDMov.InsertaMovRef lnMovNro, rsDocs!nMovNro, , nTipoPago
             '*************************************************************
             If nMontoPago >= rsDocs!Importe Then
             
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeB, 0, False
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeS, 0, False
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeRH, 0, False
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtacontDebeRHJ, 0, False
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeSegu, 0, False
                oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebePagoVarios, 0, False
                nMontoPago = nMontoPago - rsDocs!Importe
             Else
                If pnImporteB <> 0 Then
                    oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebeB, rsDocs!Importe - nMontoPago
                Else
                   If pnImporteS <> 0 Then
                      'oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebeS, rsDocs!Importe - nMontoPago '*** PEAC 20110301
                      oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebeS, rsDocs!Importe - rsDocs!VALORSOLES
                   Else
                      If pnImporteRH <> 0 Then
                        oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebeRH, rsDocs!Importe - nMontoPago
                      Else
                        If pnImporteRHJ <> 0 Then
                           oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtacontDebeRHJ, rsDocs!Importe - nMontoPago
                        Else
                            If pnImportePagoVarios <> 0 Then
                               oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebePagoVarios, rsDocs!Importe - nMontoPago
                            Else
                               oDMov.InsertaMovPendientesRend rsDocs!nMovNro, psCtaContDebeSegu, rsDocs!Importe - nMontoPago
                            End If
                        End If
                      End If
                   End If
                End If
                'nMontoPago = 0 '*** PEAC 20110301
             End If
          End If
          rsDocs.MoveNext
       Loop
    'End If

End If
End If

'Fin Agregado

'--Calculo del ITF
'--  JOHN
If bAfectoITF = True Then 'bit si esta afecto al ITF el pago a Proveedor
    'Dim lnMontoITF As Currency
    Dim lnMontoITF As Double '*** PEAC 20110331
    If lsCtaITFD <> "" And lsCtaITFH <> "" And nTasaITF > 0 Then
        'Modificado PASI20150415
        'lnMontoITF = fgTruncar((pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu + pnImportePagoVarios - lnRetAct) * nTasaITF, 2)
        lnMontoITF = Round(DameMontoITF(pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu + pnImportePagoVarios - lnRetAct), 2)
        'end PASI
        If lnMontoITF > 0 Then
            lnMovItem = lnMovItem + 1
            'ALPA 20090317*********************************************
            'oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFD, lnMontoITF
            oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFD, lnMontoITF, nTipoPago
            '**********************************************************
            lnMovItem = lnMovItem + 1
            'ALPA 20090317*********************************************
            'oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFH, lnMontoITF * -1
            oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFH, lnMontoITF * -1, nTipoPago
            '**********************************************************
        End If
    End If
End If
'----------

'EJVG20131209 *** Incluimos la Penalidad
Set rsPenalidad = New ADODB.Recordset
Set rsPenalidad = oDMov.ListaItemsPenalidad(pnMovNroProv)
If Not rsPenalidad.EOF Then
    Do While Not rsPenalidad.EOF
        lnMovItem = lnMovItem + 1
        lnPenalidad = lnPenalidad + rsPenalidad!nMovImporte
        oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), rsPenalidad!cCtaContCod, rsPenalidad!nMovImporte, nTipoPago
        rsPenalidad.MoveNext
    Loop
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaContPenalidadH, lnPenalidad * -1, nTipoPago
End If
'END EJVG *******

If psOpecod <> "402581" And psOpecod <> "402582" Then
    If Mid(psOpecod, 3, 1) = 2 Then
        If lbBitReten Then
            'ALPA 20090430****************************
            'oDMov.GeneraMovME lnMovNro, psMovNro, gnTipCambioPonderado
            oDMov.GeneraMovME lnMovNro, psMovNro, gnTipCambioPonderado, , , nTipoPago
            '*****************************************
        Else
            'ALPA 20090430****************************
            'oDMov.GeneraMovME lnMovNro, psMovNro
            oDMov.GeneraMovME lnMovNro, psMovNro, , , , nTipoPago
            '********************************
        End If
    End If
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"

If psDocTpo = TpoDocNotaAbono Or psDocTpo = TpoDocNotaCargo Then
    oDMov.InsertaRegDocCuenta psDocTpo, psDocNro, psCuentaAho
    lbGraboNegocio = False
    If pbGrabaOpeNegocio Then
        Dim oCapta As NCapMovimientos
        Dim lnSaldo As Double
        Set oCapta = New NCapMovimientos
        Dim oCon As New DConecta
        Dim oDis As New NRHProcesosCierre
        If gbBitCentral Then
           lsMovNroNegocio = oDMov.GeneraMovNro(, , , psMovNro)
           lbGraboNegocio = True
           
           'CTI2 20190502 Comentado
           'lnSaldo = oCapta.CapAbonoCuentaAho(psCuentaAho, pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu + pnImportePagoVarios - lnRetAct, gCGPagProvAbonoCent, lsMovNroNegocio, psMovDesc, TpoDocNotaAbono, psDocNro, , , , , , , , , , True, False, , , oDMov.GetConexion, False, bAfectoITF, CDbl(lnMontoITF))
           
           'CTI2 20190502 Add
           lnSaldo = oDMov.CapAbonoCuentaAho(psCuentaAho, pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu + pnImportePagoVarios - lnRetAct, gCGPagProvAbonoCent, lsMovNroNegocio, psMovDesc, TpoDocNotaAbono, psDocNro, , , , , , , , , , True, False, , , , False, bAfectoITF, CDbl(lnMontoITF))
        Else
            If oCon.AbreConexion(Left(psCuentaAho, 2)) Then
                oCon.BeginTrans
                lbGraboNegocio = True
                lnSaldo = oDis.Abono(psCuentaAho, pnImporteB + pnImporteS + pnImporteRH + pnImporteRHJ + pnImporteSegu + pnImportePagoVarios - lnRetAct, gCGPagProvAbonoDist, gCGPagProvAbonoDist, "112" & Left(psCuentaAho, 2), Right(psMovNro, 4), psDocNro, "PAGO : " & psMovDesc, oCon, GetFechaMov(psMovNro, True))
                oCon.CommitTrans
            End If
            oCon.CierraConexion
            Set oCon = Nothing
            Set oDis = Nothing
        End If
    End If
End If

If pnMovNroProv > 0 Then
    oDMov.ActualizaEstadoCuotasCronogramaxPagoProveedor pnMovNroProv
    oDMov.ActualizaEstadoCuotasCronogramaxPagoProveedorNew pnMovNroProv 'PASIERS0772014
    oDMov.ActualizaEstadoBienesxPagoProveedor pnMovNroProv 'PASIERS0772014
    
End If
If Not oDMov.ExisteMovRef(lnMovNro, pnMovNroProv) Then 'EJVG20140407 Verifica se inserte Relacion Pago con Provisión
    oDMov.InsertaMovRef lnMovNro, pnMovNroProv, , nTipoPago
End If

oDMov.CommitTrans
lbTrans = False
GrabaPagoProveedor_NEW = 1

Set oDMov = Nothing

Exit Function
GrabaPagoProveedorErr:
   lsMsgErr = Err.Description
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    If lbGraboNegocio Then
        lsMsgErr = lsMsgErr & Chr(10) & Chr(10) & "EL SISTEMA HIZO EL ABONO EN LA CUENTA DEL CLIENTE. POR FAVOR EXTORNAR ESTA OPERACION"
    End If
    Set oDMov = Nothing
    Err.Raise vbObjectError + 100, "GrabaPagoProveedor", lsMsgErr
End Function
'END EJVG *******
'FRHU20140118 RQ13825
Public Sub GrabarProyeccionAhorroAge(ByVal pnAnio As String, ByVal rsFCP As ADODB.Recordset)
    Dim oFlujoCaja As New DCajaGeneral
    Dim lnNuevoCod As Integer
    Dim lbTrans As Boolean
    lbTrans = False
    
    On Error GoTo MsgError
    oFlujoCaja.dBeginTrans
    
    Do While Not rsFCP.EOF
        Call oFlujoCaja.InsertaProyeccionAhorroAge(pnAnio, "01", rsFCP![AgeCod], rsFCP![Enero])
        Call oFlujoCaja.InsertaProyeccionAhorroAge(pnAnio, "02", rsFCP![AgeCod], rsFCP![Febrero])
        Call oFlujoCaja.InsertaProyeccionAhorroAge(pnAnio, "03", rsFCP![AgeCod], rsFCP![Marzo])
        Call oFlujoCaja.InsertaProyeccionAhorroAge(pnAnio, "04", rsFCP![AgeCod], rsFCP![Abril])
        Call oFlujoCaja.InsertaProyeccionAhorroAge(pnAnio, "05", rsFCP![AgeCod], rsFCP![Mayo])
        Call oFlujoCaja.InsertaProyeccionAhorroAge(pnAnio, "06", rsFCP![AgeCod], rsFCP![Junio])
        Call oFlujoCaja.InsertaProyeccionAhorroAge(pnAnio, "07", rsFCP![AgeCod], rsFCP![Julio])
        Call oFlujoCaja.InsertaProyeccionAhorroAge(pnAnio, "08", rsFCP![AgeCod], rsFCP![Agosto])
        Call oFlujoCaja.InsertaProyeccionAhorroAge(pnAnio, "09", rsFCP![AgeCod], rsFCP![Septiembre])
        Call oFlujoCaja.InsertaProyeccionAhorroAge(pnAnio, "10", rsFCP![AgeCod], rsFCP![Octubre])
        Call oFlujoCaja.InsertaProyeccionAhorroAge(pnAnio, "11", rsFCP![AgeCod], rsFCP![Noviembre])
        Call oFlujoCaja.InsertaProyeccionAhorroAge(pnAnio, "12", rsFCP![AgeCod], rsFCP![Diciembre])
        rsFCP.MoveNext
    Loop
    
    oFlujoCaja.dCommitTrans
    lbTrans = True
    
    Set oFlujoCaja = Nothing
    Exit Sub
MsgError:
    If Not lbTrans Then
        oFlujoCaja.dRollbackTrans
    End If
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Error"
End Sub
'FIN FRHU20140118 RQ13825
'EJVG20140512 ***
Public Function GrabarRemesaIFiAAgencia(ByVal pdFecha As Date, ByVal psAgeCod As String, ByVal psUser As String, ByVal psOpecod As String, _
                                        ByVal psIFTpo As String, ByVal psPersCodIf As String, ByVal psIFCtaCod As String, _
                                        ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal psAreaAgeCodOrig As String, ByVal psAreaAgeCodDest As String, _
                                        ByVal psGlosa As String, ByVal pnImporte As Currency, ByVal pbPropioTransp As Boolean, ByVal psPersCodTransp As String, _
                                        ByVal psCtaContCodD As String, ByVal psCtaContCodH As String, ByRef psMovNro As String) As Boolean
    On Error GoTo ErrGrabarRemesaIFiAAgencia
    Dim oMov As New DMov
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lnMovItem As Long
    Dim bTrans As Boolean
    
    oMov.BeginTrans
    bTrans = True
    lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, psUser)
    oMov.InsertaMov lsMovNro, psOpecod, Left(psGlosa, 300), gMovEstContabMovContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(lsMovNro)
    lnMovItem = 1
    oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContCodD, pnImporte
    lnMovItem = lnMovItem + 1
    oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContCodH, pnImporte * -1
    oMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
    oMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, psPersCodIf, psIFTpo, psIFCtaCod
    'oMov.InsertaMovHabilitaciones lnMovNro, "", Mid(psAreaAgeCodOrig, 1, 3), Mid(psAreaAgeCodOrig, 4, 2), Mid(psAreaAgeCodDest, 1, 3), Mid(psAreaAgeCodDest, 4, 2), 0, pnImporte, Mid(psOpeCod, 3, 1)
    oMov.InsertaMovHabilitaciones lnMovNro, "", Mid(psAreaAgeCodOrig, 1, 3), "", Mid(psAreaAgeCodDest, 1, 3), Mid(psAreaAgeCodDest, 4, 2), 0, pnImporte, Mid(psOpecod, 3, 1)
    oMov.InsertaMovHabilitacionTransp lnMovNro, pbPropioTransp, psPersCodTransp
    If psNroDoc <> "" Then
        oMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, Format(pdFecha, gsFormatoFecha)
    End If
    
    If Mid(psOpecod, 3, 1) = gMonedaExtranjera Then
        oMov.GeneraMovME lnMovNro, lsMovNro
    End If
    oMov.ActualizaSaldoMovimiento lsMovNro, "+"
    oMov.CommitTrans
    bTrans = False
    Set oMov = Nothing
    psMovNro = lsMovNro
    GrabarRemesaIFiAAgencia = True
    Exit Function
ErrGrabarRemesaIFiAAgencia:
    GrabarRemesaIFiAAgencia = False
    If bTrans Then
        oMov.RollbackTrans
        Set oMov = Nothing
    End If
    Err.Raise vbObjectError + 100, "GrabarRemesaIFiAAgencia", "Error al Grabar Remesa de Inst Financ. a Agencia"
End Function
Public Function ListaSolicitudHabRemesaxAprobacion(ByVal pnMoneda As Moneda, ByVal pdFecha As Date) As ADODB.Recordset
    Dim sql As String
    On Error GoTo ErrListaSolicitudHabRemesaxAprobacion
    sql = "EXEC stp_sel_ERS0252014_ListaSolicitudHabRemesaxAprobar " & pnMoneda & ",'" & Format(pdFecha, "yyyymmdd") & "'"
    Set ListaSolicitudHabRemesaxAprobacion = oConect.CargaRecordSet(sql)
    Exit Function
ErrListaSolicitudHabRemesaxAprobacion:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function AprobarSolicitudHabRemesa(ByRef psIDs As String, ByVal pbAprobar As Boolean, ByVal psMovNro As String) As Boolean
    Dim sql As String
    On Error GoTo ErrListaSolicitudHabRemesaxAprobacion
    sql = "EXEC stp_upd_ERS0252014_AprobarSolicitudHabRemesa '" & psIDs & "'," & IIf(pbAprobar, 1, 0) & ",'" & psMovNro & "'"
    oConect.Ejecutar sql
    AprobarSolicitudHabRemesa = True
    Exit Function
ErrListaSolicitudHabRemesaxAprobacion:
    AprobarSolicitudHabRemesa = False
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function ListaRemesaIFiToAgenciaxExtorno(ByVal pdFecha As Date, ByVal pnMoneda As Moneda) As ADODB.Recordset
    Dim sql As String
    On Error GoTo ErrListaRemesaIFiToAgenciaxExtorno
    sql = "EXEC stp_sel_ERS0252014_ListaHabilitacionRemesaIFixExtorno '" & Format(pdFecha, "yyyymmdd") & "'," & pnMoneda
    Set ListaRemesaIFiToAgenciaxExtorno = oConect.CargaRecordSet(sql)
    Exit Function
ErrListaRemesaIFiToAgenciaxExtorno:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function ExtornaRemesaIFiAAgencia(ByVal pDatos As Variant, ByVal pdFecha As Date, ByVal psAgeCod As String, ByVal psUser As String, ByVal psOpecod As String, ByVal psGlosa As String) As Boolean
    On Error GoTo ErrExtornaRemesaIFiAAgencia
    Dim oMov As New DMov
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim fila As Integer
    Dim bTrans As Boolean
    
    oMov.BeginTrans
    bTrans = True
    lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, psUser)
    oMov.InsertaMov lsMovNro, psOpecod, psGlosa, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(lsMovNro)
    
    For fila = 1 To UBound(pDatos, 2)
        oMov.ActualizaMov pDatos(2, fila), , , gMovFlagExtornado
        oMov.InsertaMovRef lnMovNro, pDatos(2, fila)
        oMov.ActualizaSaldoMovimiento CStr(pDatos(3, fila)), "-"
    Next
    
    oMov.CommitTrans
    bTrans = False
    Set oMov = Nothing
    ExtornaRemesaIFiAAgencia = True
    Exit Function
ErrExtornaRemesaIFiAAgencia:
    ExtornaRemesaIFiAAgencia = False
    If bTrans Then
        oMov.RollbackTrans
        Set oMov = Nothing
    End If
    Err.Raise vbObjectError + 100, "ExtornaRemesaIFiAAgencia", "Error al Grabar Remesa de Inst Financ. a Agencia"
End Function
'END EJVG *******
'FRHU 20140814: ANEXO 3-ERS068-2014
Public Function GetCajaBancosOperacionesSegTarjeta(ByVal psOpecod As String, ByVal pdFechaDel As Date, ByVal pdFechaAl As Date) As ADODB.Recordset
On Error GoTo GetCajaBancosOperacionesSegTarjetaErr
    sql = "exec stp_sel_getCajaBancosOperacionesSegTarjeta '" & psOpecod & "','" & Format(pdFechaDel, gsFormatoMovFecha) & "','" & Format(pdFechaAl, gsFormatoMovFecha) & "'"
    Set GetCajaBancosOperacionesSegTarjeta = oConect.CargaRecordSet(sql)
    Exit Function
GetCajaBancosOperacionesSegTarjetaErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub GrabaExtornoMovDepositoSegTarjeta(ByVal psMovNroExt As String, ByVal pnMovNroAnt As Long)
On Error GoTo ErrorGrabaExtornoMov
    Dim oMov As DMov
    Dim lbTrans As Boolean
    
    
    Set oMov = New DMov
    oMov.BeginTrans
    lbTrans = True

    Call oMov.ExtornarMovimientoDepositoSegTarjetaReferencia(psMovNroExt, pnMovNroAnt)

    oMov.CommitTrans
    lbTrans = False
    Set oMov = Nothing
    Exit Sub
ErrorGrabaExtornoMov:
    lsMsgErr = Err.Description
    If lbTrans Then oMov.RollbackTrans
    Err.Raise Err.Number, "GrabaExtornoMov", lsMsgErr
End Sub
'FIN FRHU 20140814
'PASIERS1242014
Public Function GrabaDevolucionPagoProvSUNAT(pdFecsis As Date, psCodAge As String, psUser As String, psOpecod As String, psMovDesc As String, psCtaContDebe As String, psCtaContHaber As String, ByVal psCtaContHaber2 As String, psCtaCajaGeneral As String, pnImporte As Currency, _
                                                                        pnImporteTotal As Currency, psPersCod As String, psTpoIFD As String, psPersCodIFD As String, psCtaIFD As String, psTpoIFH As String, psPersCodIFH As String, psCtaIFH As String, psTpoIFH2 As String, psPersCodIFH2 As String, psCtaIFH2 As String, _
                                                                        psDocTpo As String, psDocNro As String, psFechaDoc As String, psDocVoucher As String, psDocResSUNAT As String, psFechaDocResSUNAT As String, psDocTpoH2 As String, psDocNroH2 As String, psFechaDocH2 As String, psDocVoucherH2 As String, _
                                                                        psCuentaAho As String, Optional nTipoPago As Integer = 0, Optional pbGrabaOpeNegocio As Boolean = False, Optional lsCtaITFD As String, Optional lsCtaITFH As String, Optional nTasaITF As Double = 0, Optional bAfectoITF As Boolean = True, _
                                                                        Optional pnMovNroPagoSUNAT As Long = 0, Optional psMovDesc2 As String = "") As String
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim lnMovNro As Long
Dim lbTrans As Boolean
Dim lsMovNro As String
Dim oDMov As DMov
Set oDMov = New DMov

Dim oConst As NConstSistemas
Set oConst = New NConstSistemas
Dim oImp As DImpuesto
Set oImp = New DImpuesto
Dim oFun As NContFunciones
Set oFun = New NContFunciones

 Dim oImpr As NContImprimir
Set oImpr = New NContImprimir

Dim lsMovNroNegocio As String
Dim lbGraboNegocio As Boolean
Dim lsNroAsiento() As String
Dim lsCadenaAsiento As String

Dim I As Integer

On Error GoTo ErrGrabaDevolucionPagoProvSUNAT

oDMov.BeginTrans
lbTrans = True
    lnMovItem = 0
    lnMovOrden = 0
    ReDim Preserve lsNroAsiento(1 To 1, 2)
    lsMovNro = oFun.GeneraMovNro(pdFecsis, Right(psCodAge, 2), psUser)
    lsNroAsiento(1, 1) = lsMovNro
    GetTipCambio GetFechaMov(lsMovNro, True)
    oDMov.InsertaMov lsMovNro, "421124", psMovDesc, gMovEstContabMovContable, gMovFlagVigente
    lnMovNro = oDMov.GetnMovNro(lsMovNro)
    oDMov.InsertaMovCont lnMovNro, pnImporteTotal, 0, "0"
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebe, pnImporteTotal, nTipoPago
    lnMovOrden = lnMovOrden + 1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCodIFD, psTpoIFD, psCtaIFD
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaCajaGeneral, pnImporteTotal * -1, nTipoPago
    If Not oDMov.ExisteMovRef(lnMovNro, pnMovNroPagoSUNAT) Then
        oDMov.InsertaMovRef lnMovNro, pnMovNroPagoSUNAT, , nTipoPago
    End If
    
    lnMovItem = 0
    lnMovOrden = 0
    lsMovNro = oDMov.GeneraMovNro(, , , lsMovNro)
    lsNroAsiento(1, 2) = lsMovNro
    GetTipCambio GetFechaMov(lsMovNro, True)
    oDMov.InsertaMov lsMovNro, "421124", psMovDesc, gMovEstContabMovContable, gMovFlagVigente
    lnMovNro = oDMov.GetnMovNro(lsMovNro)
    oDMov.InsertaMovCont lnMovNro, pnImporteTotal, "0", "0"
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber, pnImporte * -1, nTipoPago
    lnMovOrden = lnMovOrden + 1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCodIFH, psTpoIFH, psCtaIFH
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaCajaGeneral, pnImporte, nTipoPago
    If Not oDMov.ExisteMovRef(lnMovNro, pnMovNroPagoSUNAT) Then
        oDMov.InsertaMovRef lnMovNro, pnMovNroPagoSUNAT, , nTipoPago
    End If
   If psDocNro <> "" Then
        oDMov.InsertaMovDoc lnMovNro, psDocTpo, psDocNro, psFechaDoc
   End If
   If psDocVoucher <> "" Then
        oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocVoucher, GetFechaMov(lsMovNro, False)
   End If
   If psDocResSUNAT <> "" Then
        oDMov.InsertaMovDoc lnMovNro, 154, psDocResSUNAT, psFechaDocResSUNAT
   End If
   oDMov.InsertaMovGasto lnMovNro, psPersCod, "0"
   If bAfectoITF = True Then
        Dim lnMontoITF As Double
            If lsCtaITFD <> "" And lsCtaITFH <> "" And nTasaITF > 0 Then
                lnMontoITF = fgTruncar((pnImporte) * nTasaITF, 2)
                If lnMontoITF > 0 Then
                    lnMovItem = lnMovItem + 1
                    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFD, lnMontoITF, nTipoPago
                    lnMovItem = lnMovItem + 1
                    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFH, lnMontoITF * -1, nTipoPago
                End If
            End If
    End If
    
   If pnImporte < pnImporteTotal Then
        lnMovItem = 0
        lnMovOrden = 0
         ReDim Preserve lsNroAsiento(1 To 1, 3)
        lsMovNro = oDMov.GeneraMovNro(, , , lsMovNro)
        lsNroAsiento(1, 3) = lsMovNro
        oDMov.InsertaMov lsMovNro, "421124", psMovDesc2, gMovEstContabMovContable, gMovFlagVigente
        lnMovNro = oDMov.GetnMovNro(lsMovNro)
        oDMov.InsertaMovCont lnMovNro, (pnImporteTotal - pnImporte), "0", "0"
        lnMovItem = lnMovItem + 1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber2, (pnImporteTotal - pnImporte) * -1, nTipoPago
        lnMovOrden = lnMovOrden + 1
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCodIFH2, psTpoIFH2, psCtaIFH2
        lnMovItem = lnMovItem + 1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaCajaGeneral, (pnImporteTotal - pnImporte), nTipoPago
        If Not oDMov.ExisteMovRef(lnMovNro, pnMovNroPagoSUNAT) Then
            oDMov.InsertaMovRef lnMovNro, pnMovNroPagoSUNAT, , nTipoPago
        End If
        If psDocNroH2 <> "" Then
            oDMov.InsertaMovDoc lnMovNro, psDocTpoH2, psDocNroH2, psFechaDocH2
        End If
        If psDocVoucherH2 <> "" Then
            oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocVoucherH2, GetFechaMov(lsMovNro, False)
        End If
    End If
    oDMov.CommitTrans
    lbTrans = False
    For I = 1 To UBound(lsNroAsiento, 2)
        oDMov.ActualizaSaldoMovimiento lsNroAsiento(1, I), "+"
        lsCadenaAsiento = lsCadenaAsiento + oImpr.ImprimeAsientoContable(lsNroAsiento(1, I), 66, 79)
    Next
    GrabaDevolucionPagoProvSUNAT = lsCadenaAsiento
    Set oDMov = Nothing
    Exit Function
ErrGrabaDevolucionPagoProvSUNAT:
   lsMsgErr = Err.Description
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    If lbGraboNegocio Then
        lsMsgErr = lsMsgErr & Chr(10) & Chr(10) & "EL SISTEMA HIZO EL ABONO EN LA CUENTA DEL CLIENTE. POR FAVOR EXTORNAR ESTA OPERACION"
    End If
    Set oDMov = Nothing
    Err.Raise vbObjectError + 100, "GrabaDevolucionPagoProvSUNAT", lsMsgErr
End Function
Public Function GrabaDevolucionPagoProvSUNATxAbonoCta(psMovNro As String, psOpecod As String, psMovDesc As String, psCtaContDebe As String, psCtaContHaber As String, ByVal psCtaContHaber2 As String, psCtaCajaGeneral As String, pnImporte As Currency, _
                                                                        pnImporteTotal As Currency, psPersCod As String, psTpoIFD As String, psPersCodIFD As String, psCtaIFD As String, psTpoIFH As String, psPersCodIFH As String, psCtaIFH As String, psTpoIFH2 As String, psPersCodIFH2 As String, psCtaIFH2 As String, _
                                                                        psDocTpo As String, psDocNro As String, psFechaDoc As String, psDocVoucher As String, psDocResSUNAT As String, psFechaDocResSUNAT As String, psDocTpoH2 As String, psDocNroH2 As String, psFechaDocH2 As String, psDocVoucherH2 As String, _
                                                                        psCuentaAho As String, Optional nTipoPago As Integer = 0, Optional pbGrabaOpeNegocio As Boolean = False, Optional lsCtaITFD As String, Optional lsCtaITFH As String, Optional nMontoITF As Currency = 0, Optional bAfectoITF As Boolean = True, _
                                                                        Optional pnMovNroPagoSUNAT As Long = 0, Optional psMovDesc2 As String = "") As String 'PASI agrego nMontoITF
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim lnMovNro As Long
Dim lbTrans As Boolean
Dim lsMovNro As String
Dim oDMov As DMov
Set oDMov = New DMov

Dim oConst As NConstSistemas
Set oConst = New NConstSistemas
Dim oImp As DImpuesto
Set oImp = New DImpuesto
Dim oFun As NContFunciones
Set oFun = New NContFunciones

 Dim oImpr As NContImprimir
Set oImpr = New NContImprimir


Dim lsMovNroNegocio As String
Dim lbGraboNegocio As Boolean
Dim lsNroAsiento() As String
Dim lsCadenaAsiento As String

Dim I As Integer

Dim lnMontoITF As Double

On Error GoTo ErrGrabaDevolucionPagoProvSUNAT

oDMov.BeginTrans
lbTrans = True
    lnMovItem = 0
    lnMovOrden = 0
    ReDim Preserve lsNroAsiento(1 To 1, 2)
    lsNroAsiento(1, 1) = psMovNro
    oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
    lnMovNro = oDMov.GetnMovNro(psMovNro)
    oDMov.InsertaMovCont lnMovNro, pnImporteTotal, 0, "0"
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebe, pnImporteTotal, nTipoPago
    lnMovOrden = lnMovOrden + 1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCodIFD, psTpoIFD, psCtaIFD
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaCajaGeneral, pnImporteTotal * -1, nTipoPago
    If Not oDMov.ExisteMovRef(lnMovNro, pnMovNroPagoSUNAT) Then
        oDMov.InsertaMovRef lnMovNro, pnMovNroPagoSUNAT, , nTipoPago
    End If
    lnMovItem = 0
    lnMovOrden = 0
    lsMovNro = oDMov.GeneraMovNro(, , , psMovNro)
    lsNroAsiento(1, 2) = lsMovNro
    GetTipCambio GetFechaMov(lsMovNro, True)
    oDMov.InsertaMov lsMovNro, "421124", psMovDesc, gMovEstContabMovContable, gMovFlagVigente
    lnMovNro = oDMov.GetnMovNro(lsMovNro)
    oDMov.InsertaMovCont lnMovNro, pnImporte, 0, "0"
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber, pnImporte * -1, nTipoPago
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaCajaGeneral, pnImporte, nTipoPago
    If Not oDMov.ExisteMovRef(lnMovNro, pnMovNroPagoSUNAT) Then
        oDMov.InsertaMovRef lnMovNro, pnMovNroPagoSUNAT, , nTipoPago
    End If
   If psDocNro <> "" Then
        oDMov.InsertaMovDoc lnMovNro, psDocTpo, psDocNro, psFechaDoc
   End If
   If psDocVoucher <> "" Then
        oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocVoucher, GetFechaMov(psMovNro, False)
   End If
   If psDocResSUNAT <> "" Then
        oDMov.InsertaMovDoc lnMovNro, 154, psDocResSUNAT, psFechaDocResSUNAT
   End If
   oDMov.InsertaMovGasto lnMovNro, psPersCod, "0"
    If bAfectoITF = True Then
        
            If lsCtaITFD <> "" And lsCtaITFH <> "" Then
                'lnMontoITF = fgTruncar((pnImporte) * nTasaITF, 2) 'Comentado PASI20150414
                lnMontoITF = nMontoITF ' PASI20150414
                If lnMontoITF > 0 Then
                    lnMovItem = lnMovItem + 1
                    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFD, lnMontoITF, nTipoPago
                    lnMovItem = lnMovItem + 1
                    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFH, lnMontoITF * -1, nTipoPago
                End If
            End If
    End If
    If pnImporte < pnImporteTotal Then
        lnMovItem = 0
        lnMovOrden = 0
         ReDim Preserve lsNroAsiento(1 To 1, 3)
        lsMovNro = oDMov.GeneraMovNro(, , , lsMovNro)
        lsNroAsiento(1, 3) = lsMovNro
        oDMov.InsertaMov lsMovNro, "421124", psMovDesc2, gMovEstContabMovContable, gMovFlagVigente
        lnMovNro = oDMov.GetnMovNro(lsMovNro)
        oDMov.InsertaMovCont lnMovNro, (pnImporteTotal - pnImporte), "0", "0"
        lnMovItem = lnMovItem + 1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber2, (pnImporteTotal - pnImporte) * -1, nTipoPago
        lnMovOrden = lnMovOrden + 1
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCodIFH2, psTpoIFH2, psCtaIFH2
        lnMovItem = lnMovItem + 1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaCajaGeneral, (pnImporteTotal - pnImporte), nTipoPago
        If Not oDMov.ExisteMovRef(lnMovNro, pnMovNroPagoSUNAT) Then
            oDMov.InsertaMovRef lnMovNro, pnMovNroPagoSUNAT, , nTipoPago
        End If
        If psDocNroH2 <> "" Then
            oDMov.InsertaMovDoc lnMovNro, psDocTpoH2, psDocNroH2, psFechaDocH2
        End If
        If psDocVoucherH2 <> "" Then
            oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocVoucherH2, GetFechaMov(lsMovNro, False)
        End If
    End If
    If psDocTpo = TpoDocNotaAbono Or psDocTpo = TpoDocNotaCargo Then
        oDMov.InsertaRegDocCuenta psDocTpo, psDocNro, psCuentaAho
        lbGraboNegocio = False
        If pbGrabaOpeNegocio Then
            Dim oCapta As NCapMovimientos
            Dim lnSaldo As Double
            Set oCapta = New NCapMovimientos
            Dim oCon As New DConecta
            Dim oDis As New NRHProcesosCierre
            lsMovNroNegocio = oDMov.GeneraMovNro(, , , lsMovNro)
            lbGraboNegocio = True
            lnSaldo = oCapta.CapAbonoCuentaAho(psCuentaAho, pnImporte, gCGPagProvAbonoCent, lsMovNroNegocio, psMovDesc, TpoDocNotaAbono, psDocNro, , , , , , , , , , True, False, , , oDMov.GetConexion, False, bAfectoITF, CDbl(lnMontoITF))
        End If
    End If
    oDMov.CommitTrans
    lbTrans = False
    For I = 1 To UBound(lsNroAsiento, 2)
        oDMov.ActualizaSaldoMovimiento lsNroAsiento(1, I), "+"
        lsCadenaAsiento = lsCadenaAsiento + oImpr.ImprimeAsientoContable(lsNroAsiento(1, I), 66, 79)
    Next
    GrabaDevolucionPagoProvSUNATxAbonoCta = lsCadenaAsiento
    Set oDMov = Nothing
    Exit Function
ErrGrabaDevolucionPagoProvSUNAT:
   lsMsgErr = Err.Description
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    If lbGraboNegocio Then
        lsMsgErr = lsMsgErr & Chr(10) & Chr(10) & "EL SISTEMA HIZO EL ABONO EN LA CUENTA DEL CLIENTE. POR FAVOR EXTORNAR ESTA OPERACION"
    End If
    Set oDMov = Nothing
    Err.Raise vbObjectError + 100, "GrabaDevolucionPagoProvSUNAT", lsMsgErr
End Function
Public Function GetDatosPagoDevolucionProveedorSunatxExtorno(pdFechaIni As Date, pdFechaFin As Date) As ADODB.Recordset
Dim sql As String
    On Error GoTo ErrGetDatosPagoDevolucionProveedorSunatxExtorno
    sql = "stp_sel_ERS1242014_GetDatosPagoDevolucionProveedorSUNATxExtorno '" & Format(pdFechaIni, "yyyyMMdd") & "','" & Format(pdFechaFin, "yyyyMMdd") & "'"
    Set GetDatosPagoDevolucionProveedorSunatxExtorno = oConect.CargaRecordSet(sql)
    Exit Function
ErrGetDatosPagoDevolucionProveedorSunatxExtorno:
    MsgBox Err.Description, vbExclamation, "Aviso"
End Function
'end PASI
'PASI20150414
Public Function DameMontoITF(ByVal pnMonto As Double) As Double
    Dim oITF As DITF
    Dim lnMonto As Double, lnITF As Double, lnRedondeoITF As Double
    
    Set oITF = New DITF
    oITF.fgITFParametros

    lnMonto = pnMonto
    lnITF = lnMonto * oITF.gnITFPorcent
    lnITF = oITF.CortaDosITF(lnITF)
    lnRedondeoITF = oITF.fgDiferenciaRedondeoITF(lnITF)
    If lnRedondeoITF > 0 Then
        lnITF = lnITF - lnRedondeoITF
    End If
    DameMontoITF = lnITF
    Set oITF = Nothing
End Function
'END PASI
'PASI20151029
Public Function ObtieneCtaxConfirmBCRP(ByVal psCtaIFCod As String) As ADODB.Recordset
    sql = "stp_sel_ObtieneCtaxConfirmBCRP '" & psCtaIFCod & "'"
    Set ObtieneCtaxConfirmBCRP = oConect.CargaRecordSet(sql)
End Function
'END PASI

