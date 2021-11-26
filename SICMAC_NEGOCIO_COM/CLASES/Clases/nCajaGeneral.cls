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
Dim lsMsgErr As String
Public Enum MotivosDepRet
    gMotDepositos = 1
    gMotRetiros = 2
End Enum

Dim SQL As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Public Function GetNroCodOp() As String
GetNroCodOp = ""
SQL = "Select Isnull(Max(cRangoCod),'0') as MaxCodRango From OpCajaGen"
Set rs = oConect.CargaRecordSet(SQL)
If Not rs.EOF And Not rs.BOF Then
    GetNroCodOp = Format(Val(rs!maxCodRango) + 1, "0000")
End If
rs.Close
Set rs = Nothing

End Function
Public Function GetEstadosOp(Optional pbFiltro As Boolean, Optional pnbEstOpDesde As CGEstadosOp = PorEmitir, Optional pnbEstOpHasta As CGEstadosOp = Extraviadas) As ADODB.Recordset
SQL = " Select  cConsDescripcion, nConsValor " _
    & " From    Constante " _
    & " Where   nCONSCOD= '" & gCGEstadosOp & "' and nCONSCOD<>nConsValor "
Set rs = oConect.CargaRecordSet(SQL)
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
SQL = "Select   cRangoCod, dFechaIng, nRangoEmiIni, nRangoEmiFin, " _
    & "         cEstadoRango, C.cConsDescripcion,  " _
    & "         Op.cMoneda, C1.cConsDescripcion as cConsDescMon, cUltActualizacion " _
    & " From    OPCajaGen op JOIN Constante C on C.nConsValor= Op.cEstadoRango " _
    & "         JOIN Constante C1 on C1.nConsValor= Op.cMoneda " _
    & " where   C.nCONSCOD= '" & gCGEstadosOp & "' and C1.nCONSCOD='" & gMoneda & "' ORDER BY OP.cRangoCod "
    
Set rs = oConect.CargaRecordSet(SQL)
Set GetOpCajaGeneral = rs
End Function
Public Function GetOpDetCajaGeneral(ByVal psRangoCod As String) As ADODB.Recordset
SQL = "Select cRangoCod, nRangoIni,nRangoFin, cEstadoDet, cConsDescripcion, cDescDet" _
    & " From OPCajaGenDet op JOIN Constante C on C.nConsValor= Op.cEstadoDet " _
    & " where C.nCONSCOD= '" & gCGEstadosOp & "' and OP.cRangoCod ='" & psRangoCod & "'"
Set rs = oConect.CargaRecordSet(SQL)
Set GetOpDetCajaGeneral = rs

End Function
Public Function VerificaOPRango(ByVal psDocNro As String, ByVal psMoneda As Moneda) As String
'PRIMERO VERIFICAMOS SI EL DOCUMENTO SE ENCUENTRA DENTRO DEL RANGO ESTABLECIDO
SQL = "     SELECT  OPG.cRangoCod,nRangoEmiIni,nRangoEmiFin,cMoneda," _
        & "         nRangoIni As RangoInIDet, nRangoFin As RangoFinDet, cEstadoDet " _
        & " From    OPCajaGen OPG LEFT JOIN OPCajaGenDet OPD ON OPD.cRangoCod= OPG.cRangoCod " _
        & " WHERE   OPG.CMONEDA ='" & psMoneda & "' AND OPG.cEstadoRango='" & PorEmitir & "' " _
        & "         AND " & psDocNro & " BETWEEN OPG.nRangoEmiIni AND OPG.nRangoEmiFin "

Set rs = oConect.CargaRecordSet(SQL)
If Not rs.EOF And Not rs.EOF Then
    'si se encuentra la op dentro del rango de las ordenes emitidas
    'buscamos ahora dentro del rango de ordenes de pagos anuladas
    'es decir en la tabla de Ordenes de pagoDetalle
    If IsNull(rs!RangoIniDet) = False Then
        SQL = " SELECT   OPG.cRangoCod,nRangoEmiIni,nRangoEmiFin,cMoneda," _
            & "         nRangoIni As RangoInIDet, nRangoFin As RangoFinDet, cEstadoDet " _
            & " From    OPCajaGen OPG LEFT JOIN OPCajaGenDet OPD ON OPD.cRangoCod= OPG.cRangoCod " _
            & " WHERE   OPG.CMONEDA ='" & psMoneda & "' AND cEstadoRango='" & PorEmitir & "' " _
            & "         AND " & psDocNro & " BETWEEN OPD.nRangoIni AND OPD.nRangoFin"
        
        Set rs = oConect.CargaRecordSet(SQL)
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
    SQL = " SELECT  OPG.cRangoCod,nRangoEmiIni,nRangoEmiFin,cMoneda," _
        & "         nRangoIni As RangoInIDet, nRangoFin As RangoFinDet, cEstadoDet " _
        & " From    OPCajaGen OPG LEFT JOIN OPCajaGenDet OPD ON OPD.cRangoCod= OPG.cRangoCod " _
        & " WHERE   OPG.CMONEDA ='" & psMoneda & "' AND cEstadoRango='" & PorEmitir & "' "
    Set rs = oConect.CargaRecordSet(SQL)
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
        ByVal rsBillete As ADODB.Recordset, ByVal rsMoneda As ADODB.Recordset, _
        ByVal psOpeCod As String, ByVal psMovDesc As String, ByVal pnImporte As Double, _
        ByVal psAreaOrig As String, ByVal psAreaDest As String, ByVal psNroDocTransp As String, _
        ByVal psFechaDocTransp As String, ByVal psNroDocSal As String, ByVal psFechaDocSal As String, _
        ByVal psCodTransp As String, ByVal pnMontoTransp As Double, ByVal nMoneda As Moneda, _
        Optional ByVal psMovNroRef As String = "", Optional ByVal psAreaCodCaja As String = "") As Integer
                                
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

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNroHab, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNroHab)


If Not rsBillete Is Nothing Then 'Billetes
    If rsBillete.State = adStateOpen Then
        If Not rsBillete.EOF And Not rsBillete.BOF Then
            Do While Not rsBillete.EOF
                oDMov.InsertaMovUserEfectivo lnMovNro, gsUsuarioBOVEDA, rsBillete("cEfectivoCod"), rsBillete("Monto")
                rsBillete.MoveNext
            Loop
        End If
        rsBillete.Close: Set rsBillete = Nothing
    End If
End If

If Not rsMoneda Is Nothing Then 'Monedas
    If rsMoneda.State = adStateOpen Then
        If Not rsMoneda.EOF And Not rsMoneda.BOF Then
            Do While Not rsMoneda.EOF
                oDMov.InsertaMovUserEfectivo lnMovNro, gsUsuarioBOVEDA, rsMoneda("cEfectivoCod"), rsMoneda("Monto")
                rsMoneda.MoveNext
            Loop
        End If
        rsMoneda.Close: Set rsMoneda = Nothing
    End If
End If

If psAreaCodCaja = "" Then
    oDMov.InsertaMovHabilitaciones lnMovNro, psCodTransp, Mid(psAreaOrig, 1, 3), Mid(psAreaOrig, 4, 2), Mid(psAreaDest, 1, 3), Mid(psAreaDest, 4, 2), pnMontoTransp, pnImporte, nMoneda
Else
    oDMov.InsertaMovHabilitaciones lnMovNro, psCodTransp, Mid(psAreaOrig, 1, 3), Mid(psAreaOrig, 4, 2), Mid(psAreaCodCaja, 1, 3), Mid(psAreaCodCaja, 4, 2), pnMontoTransp, pnImporte, nMoneda
End If

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

Public Function GetDatosHabilitacion(ByVal psMovNro As String) As ADODB.Recordset
SQL = "SELECT M.cMovNro, MH.cAreaOrig + MH.cAgeOrig + '- ' + ISNULL(AG.cAgeDescripcion,A.cAreaDescripcion ) As Origen, " _
    & "MH.cAreaDest + MH.cAgeDest + '- ' + ISNULL(AG1.cAgeDescripcion,A1.cAreaDescripcion ) As Destino, P.cPersCod, " _
    & "ISNULL(ISNULL(P.cPersNombre, T.cDescTrans ),'') As cNombreTransp, MH.nMovImporte, " _
    & "ISNULL(MH.nMontoTrans,0) As nMontoTrans,  ISNULL(C.cConsDescripcion,'') As cTipoTrans, " _
    & "ISNULL((SELECT D.cDocAbrev + '-' + MD.cDocNro FROM MovDoc MD JOIN Documento D ON D.nDocTpo = MD.nDocTpo " _
    & "WHERE MD.nMovNro = M.nMovNro AND MD.nDocTpo ='" & TpoDocCompServTranspValores & "' ),'') As DocCompTransp, " _
    & "ISNULL((SELECT D.cDocAbrev + '-' + MD.cDocNro FROM MovDoc MD JOIN Documento D ON D.nDocTpo = MD.nDocTpo " _
    & "WHERE MD.nMovNro = M.nMovNro AND MD.nDocTpo ='" & TpoDocSalidaEfectivoBoveda & "' ) ,'') As DocSalEfecBov FROM " _
    & "Mov M LEFT JOIN MovHabilitacion MH ON MH.nMovNro = M.nMovNro JOIN Areas A ON A.cAreaCod = MH.cAreaOrig " _
    & "LEFT JOIN Agencias AG ON AG.cAgeCod = MH.cAgeOrig JOIN Areas A1 ON A1.cAreaCod = MH.cAreaDest " _
    & "LEFT JOIN Agencias AG1 ON AG1.cAgeCod = MH.cAgeDest LEFT JOIN TRANSPORTE T ON T.cTranspCod = MH.cTranspCod " _
    & "LEFT JOIN Persona P ON P.cPersCod = T.cPersCod LEFT JOIN Constante C ON C.nConsValor = T.cTpoTrans AND " _
    & "C.nConsCod =" & gCGTipoTransporte & " WHERE M.cMovNro = '" & psMovNro & "' AND M.nMovFlag NOT IN " _
    & "('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "')"

Set rs = oConect.CargaRecordSet(SQL)
Set GetDatosHabilitacion = rs
End Function

Public Function GetDatosConfHabAgencias(ByVal psMovNro As String) As ADODB.Recordset
SQL = "SELECT   M.CMOVNRO, " _
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
    
Set rs = oConect.CargaRecordSet(SQL)
Set GetDatosConfHabAgencias = rs
End Function

Public Function GetHabCajaGen(ByVal psAreaOrig As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                Optional ByVal psAreaCod As String = "", Optional ByVal psAgeCod As String = "", _
                Optional ByVal pbMuestraOrigen As Boolean = True) As ADODB.Recordset

Dim lsFiltroAge As String
If psAreaCod <> "" Then
    lsFiltroAge = " AND MH.cAreaDest = '" & psAreaCod & "' AND MH.cAgeDest = '" & psAgeCod & "' "
End If

SQL = "SELECT CONVERT(Varchar(12),CONVERT(DATETIME, SUBSTRING(M.cMovNro,1,8)),103) As FechaHab, " _
    & "MH.cAreaOrig + MH.cAgeOrig + '- ' + ISNULL(AG.cAgeDescripcion,A.cAreaDescripcion ) As Origen, " _
    & "ISNULL(MH.nMovImporte,0) MontoHabilitado,  TipoTransp = '', Empresa = '', nMontoTrans = 0, " _
    & "M.cMovDesc, M.cMovNro, MH.nMoneda, MH.cAreaDest, MH.cAgeDest, M.nMovNro FROM " _
    & "Mov M JOIN MovHabilitacion MH ON MH.nMovNro = M.nMovNro JOIN Areas A ON A.cAreaCod = MH.cAreaOrig " _
    & "LEFT JOIN Agencias AG ON AG.cAgeCod = MH.cAgeOrig WHERE M.nMovFlag IN ('" & gMovFlagVigente & "') " _
    & "AND MH.cAreaOrig = '" & Mid(psAreaOrig, 1, 3) & "' AND MH.cAgeOrig = '" & Mid(psAreaOrig, 4, 2) & "' " _
    & "AND NOT EXISTS (SELECT M1.cMovNro FROM Mov M1 JOIN MovRef MR ON M1.nMovNro = MR.nMovNro " _
    & "WHERE MR.nMovNroREF = M.nMovNro AND M1.nMovFlag IN (" & gMovFlagVigente & ")) " _
    & "AND SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" _
    & Format(pdHasta, "yyyymmdd") & "' " & lsFiltroAge

Set rs = oConect.CargaRecordSet(SQL)
Set GetHabCajaGen = rs
End Function

Public Function ExtornaHabEfectivo(ByVal pdFecSis As Date, ByVal pdFechaHab As Date, ByVal psMovNroExt As String, _
                                    ByVal pnMovNroHab As Long, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                    ByVal pnImporte As Currency)
Dim oMov As DMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New DMov
On Error GoTo ExtornaHabEfectivoErr
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
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GrabaConfHabEfectivo(ByVal psMovNroConf As String, _
    ByVal psOpeCod As String, ByVal psMovDesc As String, ByVal pnImporte As Double, _
    ByVal psAreaOrig As String, ByVal psAreaDest As String, ByVal psMovNroHab As String) As Integer
            
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
oDMov.InsertaMov psMovNroConf, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNroConf)
oDMov.InsertaMovRef lnMovNro, psMovNroHab
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
Public Function GetDatosConfHabilitacion(ByVal psAreaOrig As String, _
            ByVal pdDesde As Date, ByVal pdHasta As Date, Optional ByVal psAreaCod As String = "", _
            Optional ByVal psAgeCod As String = "") As ADODB.Recordset

Dim lsFiltroAge As String

If psAreaCod <> "" Then
    lsFiltroAge = " AND MH.cAreaDest = '" & psAreaCod & "' AND MH.cAgeDest = '" & psAgeCod & "' "
End If

SQL = "SELECT CONVERT(Varchar(12),CONVERT(DATETIME, SUBSTRING(M.cMovNro,1,8)),103) As FechaHab," _
    & "ISNULL(AG.cAgeDescripcion, A.cAreaDescripcion) As Destino, ISNULL(MH.nMovImporte,0) As MontoHabilitado, " _
    & "TipoTransp = '', Empresa = '', nMontoTrans = '', M.cMovDesc, M1.cMovNro As cMovConfHab, " _
    & "MH.cAreaDest, M1.nMovNro, MH.cAgeDest, M.cMovNro " _
    & "FROM Mov M1 JOIN MovRef MR JOIN MOV M JOIN MovHabilitacion MH ON MH.nMovNro = M.nMovNro ON " _
    & "MR.nMovNroRef = M.nMovNro ON M1.nMovNro = MR.nMovNro JOIN AREAS A ON A.CAREACOD = MH.cAreaOrig " _
    & "LEFT JOIN AGENCIAS AG ON AG.cAgeCod = MH.cAgeOrig WHERE M1.nMovFlag IN (" & gMovFlagVigente & ") " _
    & "AND M.nMovFlag IN (" & gMovFlagVigente & ") " & lsFiltroAge _
    & "AND MH.cAreaOrig = '" & Mid(psAreaOrig, 1, 3) & "' AND MH.cAgeOrig = '" & Mid(psAreaOrig, 4, 2) & "' " _
    & "AND SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' " _
    
Set rs = oConect.CargaRecordSet(SQL)
Set GetDatosConfHabilitacion = rs

End Function

Public Function GrabaCompraVentaMEIF(ByVal psMovNro As String, _
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
Dim lsMsgErr As String
On Error GoTo GrabaCompraVentaMEIFErr

Set oMov = New DMov
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

oMov.GeneraMovME lnMovNro, psMovNro, pnTCBanco
If psNroDocOrig <> "" Then
    oMov.InsertaMovDoc lnMovNro, psTpoDocOrig, psNroDocOrig, Format(pdFecSis, gsFormatoFecha)
End If
If psNroDocDest <> "" Then
    oMov.InsertaMovDoc lnMovNro, psTpoDocDest, psNroDocDest, Format(pdFecSis, gsFormatoFecha)
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
Public Function GrabaCompraVentaEfectivo(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
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
Dim lsMsgErr  As String

Set oDMov = New DMov
On Error GoTo GrabaCompraVentaEfectivoErr

Set oContF = New NContFunciones

GrabaCompraVentaEfectivo = 1

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
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
                    If psOpeCod = gOpeMECompraEfect Then
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
                    If psOpeCod = gOpeMECompraEfect Then
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
                    If psOpeCod = gOpeMEVentaEfec Then
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
                    If psOpeCod = gOpeMEVentaEfec Then
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

If psOpeCod = gOpeMEVentaEfec Or psOpeCod = gOpeMECompraEfect Then
   If lnTotalD <> lnTotalS Then
      Dim sConvMES    As String
      Dim sConvMED    As String
      Dim nDif        As Currency
      Dim oCons As New NConstSistemas
      sConvMED = oCons.LeeConstSistema(gConstSistCtaConversionMEDol)
      sConvMES = oCons.LeeConstSistema(gConstSistCtaConversiónMESoles)
      Set oCons = Nothing
      lnMovItem = lnMovItem + 1
      If psOpeCod = gOpeMEVentaEfec Then
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
If psOpeCod = OpeCGOtrosOpeEfecCambme Then
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
Public Function GetMovCompraVentaME(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset
Dim lsFiltroAge As String

SQL = " SELECT  CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) as Fecha ," _
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
    & " WHERE   M.COPECOD ='" & psOpeCod & "' AND M.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "')" _
    & "         AND SUBSTRING(M.cMovnro,1,8 ) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "'  ORDER BY M.CMOVNRO "


Set rs = oConect.CargaRecordSet(SQL)
Set GetMovCompraVentaME = rs

End Function
Public Function GrabaExtornoMov(ByVal pdFecSis As Date, ByVal pdFechaMov As Date, ByVal psMovNroExt As String, _
                                  ByVal pnMovNroAnt As Long, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                  ByVal pnImporte As Currency)

On Error GoTo ErrorGrabaExtornoMov
Dim oMov As DMov
Dim lbTrans As Boolean
Set oMov = New DMov
oMov.BeginTrans
lbTrans = True
    oMov.ExtornaMovimiento psMovNroExt, pnMovNroAnt, psOpeCod, psMovDesc
oMov.CommitTrans
lbTrans = False
Set oMov = Nothing
Exit Function
ErrorGrabaExtornoMov:
    If lbTrans Then oMov.RollbackTrans
    Err.Raise Err.Number, "GrabaExtornoMov", Err.Description
End Function

Public Function GetMovCompraVentaEfectivo(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset
Dim lsFiltroAge As String

SQL = " SELECT  CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) as Fecha ," _
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
    & " WHERE   M.COPECOD ='" & psOpeCod & "' AND M.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "')" _
    & "         AND SUBSTRING(S.cMovnro,1,8 ) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' ORDER BY M.CMOVNRO "

Set rs = oConect.CargaRecordSet(SQL)
Set GetMovCompraVentaEfectivo = rs

End Function
Public Function GrabaMovimientosCtasIF(ByVal psMovNro As String, _
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
lbTrans = True
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc
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
   oMov.InsertaMovDoc lnMovNro, psTpoDocOrig, psNroDocOrig, Format(pdFecSis, gsFormatoFecha)
End If
If psNroDocDest <> "" Then
   oMov.InsertaMovDoc lnMovNro, psTpoDocDest, psNroDocDest, Format(pdFecSis, gsFormatoFecha)
End If
If Mid(psOpeCod, 3, 1) = "2" Then
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
Public Function GrabaMovEfectivo(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
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
        'oDMov.InsertaMovCtaIF lnMovNro, Mid(psObjeto, 4, 13), Mid(psObjeto, 1, 2), Mid(psObjeto, 18, 10), CGCtaIFConCapital, pnImporte * IIf(pbIngreso = False, 1, -1)
        'oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
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
    oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, Format(pdFechaDoc, gsFormatoFecha)
End If
If pnMovRef <> -1 Then
    oDMov.ActualizaMov pnMovRef, , gMovEstContabMovContable
    'oDMov.ExtornaMovCtaIF pnMovRef
    oDMov.InsertaMovRef lnMovNro, pnMovRef
End If
If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
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
Public Function GrabaMovGeneral(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
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
        'oDMov.InsertaMovCtaIF lnMovNro, Mid(psObjetoDeb, 4, 13), Mid(psObjetoDeb, 1, 2), Mid(psObjetoDeb, 18, 10), CGCtaIFConCapital, pnImporte
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
        'oDMov.InsertaMovCtaIF lnMovNro, Mid(psObjetoHab, 4, 13), Mid(psObjetoHab, 1, 2), Mid(psObjetoHab, 18, 10), CGCtaIFConCapital, pnImporte * -1
        
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
    oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, Format(pdFechaDoc, gsFormatoFecha)
End If
If psNroVoucher <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psNroVoucher, Format(pdFechaDoc, gsFormatoFecha)
End If
If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
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
Public Function GetRetirosEfectivoCaja(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset
Dim lsFiltroAge As String
SQL = " SELECT  DOCMOV.cDocAbrev, DOCMOV.cDocNro, VOUC.cDocNro AS VOUCHER, P.cPersNombre, " _
    & "         RTRIM(C1.cCtaIFDesc) + ' ' + C.cCtaIFDesc AS CUENTA, " _
    & "         ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE))  AS IMPORTE, M.NMOVNRO, M.CMOVDESC, SUBSTRING(MC.cCtaContCod,3,1) AS NMONEDA  " _
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
    & " WHERE   M.COPECOD ='" & psOpeCod & "' AND M.nMovFlag = '" & gMovFlagVigente & "' " _
    & "         AND SUBSTRING(M.cMovnro,1,8 ) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "'  " _
    & "         AND NOT EXISTS (SELECT  M1.CMOVNRO FROM  MOV M1 " _
    & "                                 JOIN MOVREF MR ON M1.nMOVNRO =MR.nMOVNRO " _
    & "                         WHERE MR.nMOVNROREF = M.nMOVNRO AND M1.nMovFlag = " & gMovFlagVigente & ") " _
    & " ORDER BY M.CMOVNRO "

Set rs = oConect.CargaRecordSet(SQL)
Set GetRetirosEfectivoCaja = rs

End Function
Public Function GrabaDepositoCheques(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
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

oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
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
'    oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, Format(pdFechaDoc, gsFormatoFecha)
'End If
If Mid(psOpeCod, 3, 1) = "2" Then
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

Public Function GrabaDevolucionCheques(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
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
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
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
                    lsFiltro = lsFiltro & oCont.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaHaber, rs!AreaCod & rs!Agecod, False)
                    lnMovItem = lnMovItem + 1
                    lnMovOrden = 0
                    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber + lsFiltro, rs!Importe * -1
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rs!cObjetoCod
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
                    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, rs!Agecod, rs!AreaCod
                    
                    oDMov.ActualizaCheque TpoDocCheque, rs![N° Cheque], rs!cPersCod, rs!cIFTpo, , , , , , gCGEstadosChqRechazado
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
If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
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
SQL = " SELECT cMotDesc, cMotCod  " _
    & " FROM MOTIVOSDEPRET WHERE cMotCod LIKE '" & pnTipoMot & "%' AND LEN(CMOTCOD)=3"

Set rs = oConect.CargaRecordSet(SQL)
Set GetMotivosDepRet = rs
End Function

Public Function GrabaDepRetBancosVarios(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                        ByVal psCtaContBanco As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                        ByVal pnDocTpo As TpoDoc, ByVal psDocNro As String, ByVal psFechaDoc As String, _
                                        ByVal psDocNroVoucher As String, ByVal rsObjHaber As ADODB.Recordset, _
                                        ByVal lnTpoObjIf As TpoObjetos, ByVal psCtaIf As String, ByVal pbDeposito As Boolean, _
                                        Optional pnMotivoNANC As Long = -1, Optional psMotObjPadre As String = "", Optional psMotObjeto As String = "", Optional ByVal psctaAho As String = "", _
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
oDMov.BeginTrans
lbTrans = True

oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
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
                Case Else
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!objeto
            End Select
            rsObjHaber.MoveNext
        Loop
    End If
End If

Select Case Val(pnDocTpo)
    Case TpoDocNotaAbono, TpoDocNotaCargo
        oDMov.ActualizaNotaAbonoCargo pnDocTpo, psDocNro, gNCNAEnMovimiento
End Select
'Optional ByVal pnDocTpoNANCAux As TpoDoc = -1,  ,  Optional ByVal psNroDocNANCAux As String = "", _
Optional pnMotivoNANCAux As Long = -1, Optional psMotObjPadreAux As String = "", Optional psMotObjetoAux As String = "", Optional ByVal psCtaAhoAux As String = ""
If psDocNro <> "" Then
    oDMov.InsertaMovDoc lnMovNro, pnDocTpo, psDocNro, Format(CDate(psFechaDoc), gsFormatoFecha)
End If
If psDocNroVoucher <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocNroVoucher, Format(CDate(psFechaDoc), gsFormatoFecha)
End If
If pnDocTpoNANCAux <> -1 And psNroDocNANCAux <> "" Then
    oDMov.ActualizaNotaAbonoCargo pnDocTpoNANCAux, psNroDocNANCAux, gNCNAEnMovimiento
    oDMov.InsertaMovDoc lnMovNro, pnDocTpoNANCAux, psNroDocNANCAux, Format(CDate(psFechaDoc), gsFormatoFecha)
End If

GrabaDepRetBancosVarios = 0
If Mid(psOpeCod, 3, 1) = 2 Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"
oDMov.CommitTrans
lbTrans = False
Set oDMov = Nothing
Exit Function
ErrorGrabaDepRetBancosVarios:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaRendicionGiroDocumento", "Error Graba Rendicion Giro de documentos"
End Function

'Public Function GrabaAperturaCtaIF(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal pdFecha As Date, ByVal psMovDesc As String, pnImporte As Currency, _
'       ByVal prsCtaApe As ADODB.Recordset, pbCreaSubCta As Boolean, psSubCtaIF As String, psSubCtaIFDesc As String, _
'       ByVal psPersCod As String, ByVal psIFTpo As String, ByVal rsBillete As ADODB.Recordset, ByVal rsMoneda As ADODB.Recordset, _
'       ByVal psNroCartaApert As String, ByVal pdFecSis As Date, _
'       ByVal psEntTransferIFCod As String, pnEntTransferIFImporte As Currency, _
'       ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pdFechaDoc As Date, ByVal psNroVoucher As String, _
'       ByVal prsCheque As ADODB.Recordset, ByVal prsOtro As ADODB.Recordset, ByVal prsObj As ADODB.Recordset, _
'       ByVal prsAdeud As ADODB.Recordset, pnGracia As Currency, pnComisionIni As Currency, psInterno As String, pnCuotaPagoK As Currency, pnNrocuotas As Integer _
'       ) As Integer
'
'Dim oContF     As NContFunciones
'Dim oDMov      As DMov
'Dim oOpe       As DOperacion
'Dim oCont      As NContFunciones
'
'Dim lnMovItem  As Long
'Dim lnMovOrden As Long
'Dim lnMovNro   As Long
'Dim lsMovNro   As String
'Dim lsSubCta   As String
'Dim lbTrans    As Boolean
'Dim lsCuenta   As String
'Dim lsFiltro   As String
'
'On Error GoTo GrabaAperturaCtaIFErr
'
'Set oDMov = New DMov
'Set oContF = New NContFunciones
'Set oOpe = New DOperacion
'Set oCont = New NContFunciones
'
'GrabaAperturaCtaIF = 1
'
'oDMov.BeginTrans
'lbTrans = True
'oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable
'lnMovNro = oDMov.GetnMovNro(psMovNro)
'oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'
''Guardamos la cuenta en el debe
''Para Adeudados hay que deducir Cuenta
'lsCuenta = oOpe.EmiteOpeCta(psOpeCod, "D", "0")
'If pbCreaSubCta Then
'    oDMov.InsertaCuentaIF psPersCod, psIFTpo, psSubCtaIF, psSubCtaIFDesc, pdFecha, "", "", 0, gEstadoCtaIFActiva, psMovNro
'End If
'lnMovItem = 0
'
'Do While Not prsCtaApe.EOF
'    If prsCtaApe!Codigo <> "" And prsCtaApe!Importe <> 0 Then
'        lnMovItem = lnMovItem + 1: lnMovOrden = 1
'        oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta, prsCtaApe!Importe
'        oDMov.InsertaCuentaIF psPersCod, psIFTpo, prsCtaApe!Codigo, prsCtaApe!Cuenta_Nro, pdFecha, "", "", nVal(prsCtaApe!Plazo), _
'                             gEstadoCtaIFRegistrada, psMovNro
'        oDMov.InsertaCuentaIFInteres psPersCod, psIFTpo, prsCtaApe!Codigo, pdFecha, nVal(prsCtaApe!Periodo), nVal(prsCtaApe!Interes), psMovNro
'        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
'        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, prsCtaApe!Codigo
'        If Not prsAdeud Is Nothing Then
'            If Not prsAdeud.State = adStateClosed Then
'                oDMov.InsertaDatosAdeudados prsAdeud, psPersCod, psIFTpo, prsCtaApe!Codigo, prsCtaApe!Importe, pnNrocuotas, pnGracia, pnComisionIni, psInterno, pnCuotaPagoK, Format(pdFecha, gsFormatoFecha), Mid(psOpeCod, 3, 1)
'            End If
'        End If
'    End If
'    prsCtaApe.MoveNext
'Loop
'
'lsCuenta = oOpe.EmiteOpeCta(psOpeCod, "H", "0")
'oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsBillete, lsCuenta, "H"
'oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsMoneda, lsCuenta, "H"
'
''Transferencia
'If psEntTransferIFCod <> "" And pnEntTransferIFImporte > 0 Then
'    lnMovItem = lnMovItem + 1: lnMovOrden = 0
'    lsCuenta = oOpe.EmiteOpeCta(psOpeCod, "H", "1", psEntTransferIFCod, Left(psEntTransferIFCod, 2))
'    oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta, pnEntTransferIFImporte * -1
'    lnMovOrden = lnMovOrden + 1
'    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
'    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psEntTransferIFCod, 4, 13), Mid(psEntTransferIFCod, 1, 2), Mid(psEntTransferIFCod, 18, 10)
'End If
'
''Cheques recibidos - Cheque de Gerencia
'If Not prsCheque Is Nothing Then
'    If prsCheque.State = adStateOpen Then
'        lsCuenta = oOpe.EmiteOpeCta(psOpeCod, "H", "2")
'        Do While Not prsCheque.EOF
'            If prsCheque!Opc = "1" Then
'                lnMovItem = lnMovItem + 1: lnMovOrden = 0
'
'                lsFiltro = oCont.GetFiltroObjetos(ObjProductosCMACT, lsCuenta, prsCheque!Cuenta, False)
'                lsFiltro = lsFiltro & oCont.GetFiltroObjetos(ObjCMACAgenciaArea, lsCuenta, prsCheque!cAreaCod & prsCheque!cAgeCod, False)
'                oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta & lsFiltro, prsCheque!Importe * -1
'
'                lnMovOrden = lnMovOrden + 1
'                oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, prsCheque!Cuenta
'                lnMovOrden = lnMovOrden + 1
'                oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
'                oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, prsCheque!cAgeCod, prsCheque!cAreaCod
'
'                oDMov.ActualizaCheque TpoDocCheque, prsCheque!NroCheque, prsCheque!cPersCod, prsCheque!cIFTpo, , , , , , gCGEstadosChqDepositado
'                oDMov.InsertaMovRef lnMovNro, prsCheque!nMovNro
'            End If
'            prsCheque.MoveNext
'        Loop
'    End If
'End If
'
''Otras Formas de Pago
'oDMov.GrabaMovOtrosCtaObj lnMovNro, lnMovItem, prsOtro, prsObj, "H"
'
'If psNroDoc <> "" Then
'    oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, Format(pdFechaDoc, gsFormatoFecha)
'End If
'If psNroVoucher <> "" Then
'    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psNroVoucher, Format(pdFechaDoc, gsFormatoFecha)
'End If
'If psNroCartaApert <> "" Then
'    oDMov.InsertaMovDoc lnMovNro, TpoDocCarta, psNroCartaApert, Format(pdFecha, gsFormatoFecha)
'End If
'
'If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
'    oDMov.GeneraMovME lnMovNro, psMovNro
'End If
'
'oDMov.ActualizaSaldoMovimiento psMovNro, "+"
'
'oDMov.CommitTrans
'lbTrans = False
'GrabaAperturaCtaIF = 0
'Set oDMov = Nothing
'Set oContF = Nothing
'
'Exit Function
'GrabaAperturaCtaIFErr:
'    lsMsgErr = Err.Description
'    If lbTrans Then
'        oDMov.RollbackTrans
'        lbTrans = False
'    End If
'    Err.Raise vbObjectError + 100, "GrabaAperturaCtaIF", lsMsgErr & " Error Graba Apertura de Ctas IF"
'End Function

'Public Function GrabaRegPagareAdeudado(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal pdFecha As Date, ByVal psMovDesc As String, pnImporte As Currency, _
'       ByVal prsCtaApe As ADODB.Recordset, pbCreaSubCta As Boolean, psSubCtaIF As String, psSubCtaIFDesc As String, _
'       ByVal psPersCod As String, ByVal psIFTpo As String, ByVal rsBillete As ADODB.Recordset, ByVal rsMoneda As ADODB.Recordset, _
'       ByVal psNroCartaApert As String, ByVal pdFecSis As Date, _
'       ByVal psEntTransferIFCod As String, pnEntTransferIFImporte As Currency, _
'       ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pdFechaDoc As Date, ByVal psNroVoucher As String, _
'       ByVal prsCheque As ADODB.Recordset, ByVal prsOtro As ADODB.Recordset, ByVal prsObj As ADODB.Recordset, _
'       ByVal prsAdeud As ADODB.Recordset, pnGracia As Currency, pnComisionIni As Currency, psInterno As String, pnCuotaPagoK As Currency, pnNrocuotas As Integer _
'       ) As Integer
'
'Dim oContF     As NContFunciones
'Dim oDMov      As DMov
'Dim oOpe       As DOperacion
'Dim oCont      As NContFunciones
'
'Dim lnMovItem  As Long
'Dim lnMovOrden As Long
'Dim lnMovNro   As Long
'Dim lsMovNro   As String
'Dim lsSubCta   As String
'Dim lbTrans    As Boolean
'Dim lsCuenta   As String
'Dim lsFiltro   As String
'
'On Error GoTo GrabaRegPagareAdeudadoErr
'
'Set oDMov = New DMov
'Set oContF = New NContFunciones
'Set oOpe = New DOperacion
'Set oCont = New NContFunciones
'
'GrabaRegPagareAdeudado = 1
'
'oDMov.BeginTrans
'lbTrans = True
'oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable
'lnMovNro = oDMov.GetnMovNro(psMovNro)
'oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'
''Guardamos la cuenta en el Haber
'lsCuenta = oOpe.EmiteOpeCta(psOpeCod, "H", "0")
'If pbCreaSubCta Then
'    oDMov.InsertaCuentaIF psPersCod, psIFTpo, psSubCtaIF, psSubCtaIFDesc, pdFecha, "", "", 0, gEstadoCtaIFActiva, psMovNro
'End If
'lnMovItem = 0
'
'Do While Not prsCtaApe.EOF
'    If prsCtaApe!Codigo <> "" And prsCtaApe!Importe <> 0 Then
'        lnMovItem = lnMovItem + 1: lnMovOrden = 1
'        oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta, prsCtaApe!Importe * -1
'        oDMov.InsertaCuentaIF psPersCod, psIFTpo, prsCtaApe!Codigo, prsCtaApe!Cuenta_Nro, pdFecha, "", "", nVal(prsCtaApe!Plazo), _
'                             gEstadoCtaIFRegistrada, psMovNro
'        oDMov.InsertaCuentaIFInteres psPersCod, psIFTpo, prsCtaApe!Codigo, pdFecha, nVal(prsCtaApe!Periodo), nVal(prsCtaApe!Interes), psMovNro
'        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
'        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, prsCtaApe!Codigo
'        If Not prsAdeud Is Nothing Then
'            If Not prsAdeud.State = adStateClosed Then
'                oDMov.InsertaDatosAdeudados prsAdeud, psPersCod, psIFTpo, prsCtaApe!Codigo, prsCtaApe!Importe, pnNrocuotas, pnGracia, pnComisionIni, psInterno, pnCuotaPagoK, Format(pdFecha, gsFormatoFecha), Mid(psOpeCod, 3, 1)
'            End If
'        End If
'    End If
'    prsCtaApe.MoveNext
'Loop
'
'lsCuenta = oOpe.EmiteOpeCta(psOpeCod, "D", "0")
'oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsBillete, lsCuenta, "D"
'oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsMoneda, lsCuenta, "D"
'
''Transferencia
'If psEntTransferIFCod <> "" And pnEntTransferIFImporte > 0 Then
'    lnMovItem = lnMovItem + 1: lnMovOrden = 0
'    lsCuenta = oOpe.EmiteOpeCta(psOpeCod, "D", "1", psEntTransferIFCod, Left(psEntTransferIFCod, 2))
'    oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta, pnEntTransferIFImporte
'    lnMovOrden = lnMovOrden + 1
'    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
'    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psEntTransferIFCod, 4, 13), Mid(psEntTransferIFCod, 1, 2), Mid(psEntTransferIFCod, 18, 10)
'End If
'
''Cheques recibidos - Cheque de Gerencia
'If Not prsCheque Is Nothing Then
'    If prsCheque.State = adStateOpen Then
'        lsCuenta = oOpe.EmiteOpeCta(psOpeCod, "D", "2")
'        Do While Not prsCheque.EOF
'            If prsCheque!Opc = "1" Then
'                lnMovItem = lnMovItem + 1: lnMovOrden = 0
'
'                lsFiltro = oCont.GetFiltroObjetos(ObjProductosCMACT, lsCuenta, prsCheque!Cuenta, False)
'                lsFiltro = lsFiltro & oCont.GetFiltroObjetos(ObjCMACAgenciaArea, lsCuenta, prsCheque!cAreaCod & prsCheque!cAgeCod, False)
'                oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta & lsFiltro, prsCheque!Importe
'
'                lnMovOrden = lnMovOrden + 1
'                oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, prsCheque!Cuenta
'                lnMovOrden = lnMovOrden + 1
'                oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
'                oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, prsCheque!cAgeCod, prsCheque!cAreaCod
'
'                oDMov.ActualizaCheque TpoDocCheque, prsCheque!NroCheque, prsCheque!cPersCod, prsCheque!cIFTpo, , , , , , gCGEstadosChqDepositado
'                oDMov.InsertaMovRef lnMovNro, prsCheque!nMovNro
'            End If
'            prsCheque.MoveNext
'        Loop
'    End If
'End If
'
''Otras Formas de Pago
'oDMov.GrabaMovOtrosCtaObj lnMovNro, lnMovItem, prsOtro, prsObj, "D"
'
'If psNroDoc <> "" Then
'    oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, Format(pdFechaDoc, gsFormatoFecha)
'End If
'If psNroVoucher <> "" Then
'    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psNroVoucher, Format(pdFechaDoc, gsFormatoFecha)
'End If
'If psNroCartaApert <> "" Then
'    oDMov.InsertaMovDoc lnMovNro, TpoDocCarta, psNroCartaApert, Format(pdFecha, gsFormatoFecha)
'End If
'
'If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
'    oDMov.GeneraMovME lnMovNro, psMovNro
'End If
'
'oDMov.ActualizaSaldoMovimiento psMovNro, "+"
'
'oDMov.CommitTrans
'lbTrans = False
'GrabaRegPagareAdeudado = 0
'Set oDMov = Nothing
'Set oContF = Nothing
'
'Exit Function
'GrabaRegPagareAdeudadoErr:
'    lsMsgErr = Err.Description
'    If lbTrans Then
'        oDMov.RollbackTrans
'        lbTrans = False
'    End If
'    Err.Raise vbObjectError + 100, "GrabaAperturaCtaIF", lsMsgErr & " Error Graba Apertura de Ctas IF"
'End Function

'Public Function GrabaPagoCuotaAdeudados(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal pdFecha As Date, ByVal psMovDesc As String, pnImporte As Currency, _
'       ByVal psCtaAdeudCod As String, pnNroCuota As Integer, ByVal rsBillete As ADODB.Recordset, ByVal rsMoneda As ADODB.Recordset, _
'       ByVal pdFecSis As Date, _
'       ByVal psEntTransferIFCod As String, pnEntTransferIFImporte As Currency, _
'       ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pdFechaDoc As Date, ByVal psNroVoucher As String, _
'       ByVal prsCheque As ADODB.Recordset, ByVal prsOtro As ADODB.Recordset, ByVal prsObj As ADODB.Recordset, _
'       ByVal prsAsiento As ADODB.Recordset, ByVal pbCancela As Boolean, ByVal pnTasaVac As Currency) As Integer
'
'Dim oContF     As NContFunciones
'Dim oDMov      As DMov
'Dim oOpe       As DOperacion
'Dim oIF        As DInstFinanc
'
'Dim lnMovItem  As Long
'Dim lnMovOrden As Long
'Dim lnMovNro   As Long
'Dim lsMovNro   As String
'Dim lsSubCta   As String
'Dim lbTrans    As Boolean
'Dim lsCuenta   As String
'Dim lsFiltro   As String
'
'Dim lnCapital  As Currency
'Dim lnInteres  As Currency
'
'On Error GoTo GrabaPagoCuotaAdeudadosErr
'
'Set oDMov = New DMov
'Set oContF = New NContFunciones
'Set oOpe = New DOperacion
'Set oIF = New DInstFinanc
'
'GrabaPagoCuotaAdeudados = 1
'
'oDMov.BeginTrans
'lbTrans = True
'oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable
'lnMovNro = oDMov.GetnMovNro(psMovNro)
'oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'
''Guardamos la cuenta en el debe
'lnMovItem = 0
'lnMovOrden = 0 'Se usara en este Bucle como Contador de Registro
'If Not prsAsiento Is Nothing Then
'    If prsAsiento.State = adStateOpen Then
'        If Not prsAsiento.EOF And Not prsAsiento.BOF Then
'            Do While Not prsAsiento.EOF
'                lnMovOrden = lnMovOrden + 1
'                If prsAsiento!Monto > 0 Then
'                    lnMovItem = lnMovItem + 1
'                    lsSubCta = ""
'                    lsSubCta = oContF.GetFiltroObjetos(ObjEntidadesFinancieras, prsAsiento!Cuenta, psCtaAdeudCod, False)
'                    If lsSubCta = "" Then
'                        lsSubCta = oIF.GetIFSubCuenta(Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2))
'                    End If
'
'                    oDMov.InsertaMovCta lnMovNro, lnMovItem, prsAsiento!Cuenta & lsSubCta, prsAsiento!Monto
'                    oDMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
'                    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10)
'                    Select Case lnMovOrden
'                        Case 1
'                            lnCapital = prsAsiento!Monto
'                        Case 2
'                            lnInteres = prsAsiento!Monto
'                    End Select
'
'                End If
'                prsAsiento.MoveNext
'            Loop
'        End If
'    End If
'End If
'
'oDMov.ActualizaAdeudadosPagoCuota Mid(psCtaAdeudCod, 4, 13), Left(psCtaAdeudCod, 2), Mid(psCtaAdeudCod, 18, 10), Format(pdFecSis, gsFormatoFecha), lnCapital, lnInteres, pnNroCuota, psMovNro, pbCancela
'If pnTasaVac > 0 And Mid(psOpeCod, 3, 1) = "1" Then
'    oDMov.InsertaMovTpoCambio lnMovNro, pnTasaVac
'End If
'
''Guardamos el efectivo
'lsCuenta = oOpe.EmiteOpeCta(psOpeCod, "H", "0")
'oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsBillete, lsCuenta, "H"
'oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsMoneda, lsCuenta, "H"
'
''Transferencia
'If psEntTransferIFCod <> "" And pnEntTransferIFImporte > 0 Then
'    lnMovItem = lnMovItem + 1: lnMovOrden = 0
'    lsCuenta = oOpe.EmiteOpeCta(psOpeCod, "H", "1", psEntTransferIFCod, Left(psEntTransferIFCod, 2))
'    oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta, pnEntTransferIFImporte * -1
'    lnMovOrden = lnMovOrden + 1
'    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
'    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psEntTransferIFCod, 4, 13), Mid(psEntTransferIFCod, 1, 2), Mid(psEntTransferIFCod, 18, 10)
'End If
'
''Cheques recibidos - Cheque de Gerencia
'If Not prsCheque Is Nothing Then
'    If prsCheque.State = adStateOpen Then
'        lsCuenta = oOpe.EmiteOpeCta(psOpeCod, "H", "2")
'        Do While Not prsCheque.EOF
'            If prsCheque!Opc = "1" Then
'                lnMovItem = lnMovItem + 1: lnMovOrden = 0
'
'                lsFiltro = oContF.GetFiltroObjetos(ObjProductosCMACT, lsCuenta, prsCheque!Cuenta, False)
'                lsFiltro = lsFiltro & oContF.GetFiltroObjetos(ObjCMACAgenciaArea, lsCuenta, prsCheque!cAreaCod & prsCheque!cAgeCod, False)
'                oDMov.InsertaMovCta lnMovNro, lnMovItem, lsCuenta & lsFiltro, prsCheque!Importe * -1
'
'                lnMovOrden = lnMovOrden + 1
'                oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, prsCheque!Cuenta
'                lnMovOrden = lnMovOrden + 1
'                oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
'                oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, prsCheque!cAgeCod, prsCheque!cAreaCod
'
'                oDMov.ActualizaCheque TpoDocCheque, prsCheque!NroCheque, prsCheque!cPersCod, prsCheque!cIFTpo, , , , , , gCGEstadosChqDepositado
'                oDMov.InsertaMovRef lnMovNro, prsCheque!nMovNro
'            End If
'            prsCheque.MoveNext
'        Loop
'    End If
'End If
'
''Otras Formas de Pago
'oDMov.GrabaMovOtrosCtaObj lnMovNro, lnMovItem, prsOtro, prsObj, "H"
'
'If psNroDoc <> "" Then
'    oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, Format(pdFechaDoc, gsFormatoFecha)
'End If
'If psNroVoucher <> "" Then
'    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psNroVoucher, Format(pdFechaDoc, gsFormatoFecha)
'End If
'
'If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
'    oDMov.GeneraMovME lnMovNro, psMovNro
'End If
'
'oDMov.ActualizaSaldoMovimiento psMovNro, "+"
'
'oDMov.CommitTrans
'lbTrans = False
'GrabaPagoCuotaAdeudados = 0
'Set oDMov = Nothing
'Set oContF = Nothing
'
'Exit Function
'GrabaPagoCuotaAdeudadosErr:
'    lsMsgErr = Err.Description
'    If lbTrans Then
'        oDMov.RollbackTrans
'        lbTrans = False
'    End If
'    Err.Raise vbObjectError + 100, "GrabaPagoCuotaAdeudados", lsMsgErr & " Error Graba Pago Adeudados"
'End Function


Public Function GrabaAperturaRegCheque(ByVal psMovNro As String, ByVal psOpeCod As String, _
                                      ByVal psMovDesc As String, _
                                      ByVal psCtaDebe As String, ByVal psObjCodProd As String, ByVal psAreaCod As String, ByVal psAgeCod As String, _
                                      ByVal psCtaHaber As String, ByVal rsObjHaber As ADODB.Recordset, _
                                      ByVal psNumCheque As String, ByVal psPersCodIF As String, ByVal psTipoIF As String, ByVal pnPlaza As ChequePlaza, _
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
oMov.BeginTrans
lbTrans = True
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovCont lnMovNro, pnImporte, 0, ""

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
oMov.InsertaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pnPlaza, psCtaCheque, pnImporte, pdFechaReg, _
                    pdFechaVal, pnEstadoCG, pnConfirmaCaja, pnMonedaCheque, , psAreaCodChq, psAgeCodChq

oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pdFechaReg, gChqEstEnValorizacion, psMovNro

If psNroDocAper <> "" Then
    oMov.InsertaMovDoc lnMovNro, pnTpoDocAper, psNroDocAper, Format(pdFechaDocApera, gsFormatoFecha)
End If
If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
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

Public Sub GeneraCuentaBancos(psCuenta As String, psSubCuenta As String, psPersCod As String, psIFTpo As String, psCtaIfCod As String, psOpeCod As String)
Dim lsMismaSubCta As Boolean
'Verificamos si Existe en CtaIFFiltro - segun Cuenta definidas en Operacion
SQL = "SELECT Distinct cif.cCtaContCod, cif.cCtaIFSubCta FROM CtaIFFiltro cif JOIN OpeCta oc ON cif.cCtaContCod = oc.cCtaContCod " _
    & "WHERE cif.cPersCod = '" & psPersCod & "' and oc.cOpeCod = '" & psOpeCod & "' and cOpeCtaDH = 'D' and cOpeCtaOpc = '1' "
Set rs = oConect.CargaRecordSet(SQL)
If Not rs.EOF Then
    psCuenta = rs!cCtaContCod
    If Val(Left(psCtaIfCod, 2)) = gTpoCtaIFCtaPF Or Val(Left(psCtaIfCod, 2)) = gTpoCtaIFCtaAdeud Then
        psSubCuenta = rs!cCtaIFSubCta
    Else
        If rs.RecordCount = 1 Then
            'Si solo hay uno => Se crearia Hijos. el Q existe seria 01 y este nuevo 02
            psSubCuenta = rs!cCtaIFSubCta & "02"
        Else
            psSubCuenta = Format(Val(rs!cCtaIFSubCta) + 1, String(Len(rs!cCtaIFSubCta), "0"))
        End If
    End If
Else
    
    'Verificamos si IF ya tiene Cuentas de Este Tipo
    SQL = "SELECT Distinct cif.cCtaContCod FROM CtaIFFiltro cif JOIN OpeCta oc ON cif.cCtaContCod = oc.cCtaContCod " _
        & "WHERE cIFTpo = '" & psIFTpo & "' and cCtaIFCod LIKE '" & Left(psCtaIfCod, 2) & "%' and oc.cOpeCod = '" & psOpeCod & "'"
    Set rs = oConect.CargaRecordSet(SQL)
    If Not rs.EOF Then
        If rs.RecordCount = 1 Then
            psCuenta = rs!cCtaContCod
        End If
    End If 'BANCO NO TIENE NINGUNA CUENTA
    If psSubCuenta = "" Then
        SQL = "SELECT cSubCtaContCod FROM InstitucionFinanc WHERE cPersCod = '" & psPersCod & "'"
        Set rs = oConect.CargaRecordSet(SQL)
        If Not rs.EOF Then
            psSubCuenta = rs!cSubCtaContCod & Left(psCtaIfCod, 2)
        End If
        
    End If
End If
'Verificamos Si existe la Cuenta Contable a Generar
SQL = "SELECT cCtaContCod FROM CtaCont WHERE cCtaContCod = '" & psCuenta & psSubCuenta & "'"
Set rs = oConect.CargaRecordSet(SQL)
If rs.EOF Then
    RSClose rs
    Err.Raise 50001, "nCajaGeneral: GeneraCuentaBancos", "Cuenta Generada [" & psCuenta & psSubCuenta & "] no Existe. Consultar con Contabilidad o Sistemas"
End If
RSClose rs
End Sub

Public Function GetApertSinConf(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset
Dim lsFiltroAge As String
Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

SQL = "Select   CONVERT(CHAR(10),CONVERT(DATETIME, SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHA," _
    & "         P.CPERSNOMBRE AS BANCO, SC.cCtaIFDesc + ' ' + C.cCtaIFDesc AS CUENTA, " _
    & "         CONVERT(CHAR(10),C.dCtaIFAper,103) AS APERTURA, ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)  AS IMPORTE,  M.NMOVNRO, M.COPECOD, C.CCTAIFCOD, M.CMOVDESC, C.cIFTpo, C.cPersCod  " _
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
    & " Where   OC.COPECOD ='" & psOpeCod & "' AND cOpeCtaDH='H' AND MC.NMOVIMPORTE>0 AND M.NMOVFLAG =" & gMovFlagVigente & "" _
    & "         AND C.cIFTpo + C.CCTAIFCOD LIKE cOpeObjFiltro  " _
    & "         AND SUBSTRING(M.cMovnro,1,8 ) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "'  " _
    & "         AND c.cCtaIFEstado = " & gEstadoCtaIFRegistrada
    
Set rs = oConect.CargaRecordSet(SQL)
Set GetApertSinConf = rs
oConect.CierraConexion
Set oConect = Nothing

End Function

Public Function GetPagareSinConf(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset
Dim lsFiltroAge As String
Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

SQL = "Select   CONVERT(CHAR(10),CONVERT(DATETIME, SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHA," _
    & "         P.CPERSNOMBRE AS BANCO, SC.cCtaIFDesc + ' ' + C.cCtaIFDesc AS CUENTA, " _
    & "         CONVERT(CHAR(10),C.dCtaIFAper,103) AS APERTURA, ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)  AS IMPORTE,  M.NMOVNRO, M.COPECOD, C.CCTAIFCOD, M.CMOVDESC, C.cIFTpo, C.cPersCod  " _
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
    & " Where   OC.COPECOD ='" & psOpeCod & "' AND cOpeCtaDH='H' AND MC.NMOVIMPORTE<0 AND M.NMOVFLAG =" & gMovFlagVigente & "" _
    & "         AND C.cIFTpo + C.CCTAIFCOD LIKE cOpeObjFiltro  " _
    & "         AND SUBSTRING(M.cMovnro,1,8 ) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "'  " _
    & "         AND c.cCtaIFEstado = " & gEstadoCtaIFRegistrada
    
Set rs = oConect.CargaRecordSet(SQL)
Set GetPagareSinConf = rs
oConect.CierraConexion
Set oConect = Nothing
End Function

Public Function CalculaSaldoAdeudadoLargoPlazo(pdFecha As Date, psPersCod As String, psIFTpo As String, psCtaIfCod As String) As Currency
On Error GoTo CalculaSaldoAdeudadoLargoPlazoErr
    CalculaSaldoAdeudadoLargoPlazo = 0
    SQL = "SELECT ISNULL(SUM(nCapital),0) nCapital FROM CtaIFCalendario " _
        & "WHERE cPersCod = '" & psPersCod & "' and cIFTpo = '" & psIFTpo & "' and cCtaIFCod = '" & psCtaIfCod & "' and " _
        & "      datediff(d,'" & Format(pdFecha, gsFormatoFecha) & "',dVencimiento) > 360 "
    Set rs = oConect.CargaRecordSet(SQL)
    If Not rs.EOF Then
        CalculaSaldoAdeudadoLargoPlazo = rs!nCapital
    End If
    RSClose rs
Exit Function
CalculaSaldoAdeudadoLargoPlazoErr:
    Err.Raise vbObjectError + 100, "CalculaSaldoAdeudadoLargoPlazo", Err.Description & " Adeudados "
End Function

Public Function GrabaConfApertura(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                ByVal psCtaDebe As String, ByVal psSubCtaDebe As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
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
Dim ldFecha As Date

Dim nImporteLP As Currency

Set oCtaIf = New NCajaCtaIF
Set oDMov = New DMov
On Error GoTo GrabaConfAperturaErr
nImporteLP = 0
If Val(Left(psCtaIfCod, 2)) = gTpoCtaIFCtaAdeud Then
   nImporteLP = CalculaSaldoAdeudadoLargoPlazo(GetFechaMov(psMovNro, True), psPersCod, psIFTpo, psCtaIfCod)
End If

GrabaConfApertura = 1
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc
lnMovNro = oDMov.GetnMovNro(psMovNro)
ldFecha = CDate(GetFechaMov(psMovNro, True))

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta de Bancos en el debe : Primer Capital
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe & psSubCtaDebe, pnImporte - nImporteLP
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIfCod

'Creamos la SubCuenta de la Inst.Financ en CtaIFFiltro
oDMov.InsertaCuentaIFFiltro psPersCod, psIFTpo, psCtaIfCod, psCtaDebe, psSubCtaDebe

'ADEUDADOS Largo Plazo
If nImporteLP > 0 Then
    lnMovOrden = 0
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), "26" & Mid(psCtaDebe, 3, 20) & psSubCtaDebe, nImporteLP
    lnMovOrden = lnMovOrden + 1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIfCod
End If

'Guardamos la Cuenta de la Pendiente que se elimina
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnImporte * -1
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIfCod

'Pasamos a activo el estado de la cuenta aperturada
oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIfCod, , , Format(ldFecha, gsFormatoFecha), Format(ldFecha, gsFormatoFecha), , gEstadoCtaIFActiva
oDMov.InsertaMovRef lnMovNro, pnMovNroRef

If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
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

Public Function GrabaInteresDev(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                ByVal psCtaDebe As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                ByVal psPersCod As String, ByVal psIFTpo As String, _
                                ByVal psCtaIfCod As String, ByVal pdCtaIFInt As Date, ByVal pnDiasTrans As Integer) As Integer

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
lnInteres = oCtaIf.GetInteresCtaIf(psPersCod, psIFTpo, psCtaIfCod)
lnInteres = lnInteres + pnImporte
Set oCtaIf = Nothing


oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc
lnMovNro = oDMov.GetnMovNro(psMovNro)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1:
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, pnImporte
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIfCod

lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnImporte * -1
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIfCod

oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIfCod, , , , Format(pdCtaIFInt, gsFormatoFecha), , , lnInteres, psMovNro
If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
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
Public Function GrabaCapitalizaRegCheque(ByVal psMovNro As String, ByVal psOpeCod As String, _
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
'oMov.InsertaMovCtaIF lnMovNro, psPersCod, pnIFTpo, psCtaIfCod, CGCtaIFConInteres, (pnImporteIntCalc + pnImporteInt) * -1, pnDiasTrans

If pnImporteCapital > 0 And psCtaHaberCapital <> "" Then
    lnMovItem = lnMovItem + 1: lnMovOrden = 0
    lnMovOrden = lnMovOrden + 1
    oMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaberCapital, pnImporteCapital * -1
    oMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, pnIFTpo, psCtaIfCod
    
    'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
    'oMov.InsertaMovCtaIF lnMovNro, psPersCod, pnIFTpo, psCtaIfCod, CGCtaIFConCapital, pnImporteCapital * -1
    'oMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                            psPersCod, pnIFTpo, psCtaIfCod
End If

'InsertaRefCap psPersCod, pnIFTpo, psCtaIfCod, lnMovNro, oMov
oMov.ActualizaCtaIF psPersCod, pnIFTpo, psCtaIfCod, , , Format(pdFechaMov, gsFormatoFecha), , , , 0, psMovNro
'grabacion dentro de la tabla de cheques
oMov.InsertaCheque TpoDocCheque, psNumCheque, psPersCodIfChq, pnTipoIFChq, pnPlaza, psCtaCheque, pnImporteTotal, pdFechaReg, _
                    pdFechaVal, pnEstadoCG, pnConfirmaCaja, pnMonedaCheque, , psAreaCodChq, psAgeCodChq
oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIfChq, pnTipoIFChq, pdFechaReg, gChqEstEnValorizacion, psMovNro

If psNumCheque <> "" Then
    oMov.InsertaMovDoc lnMovNro, TpoDocCheque, psNumCheque, Format(pdFechaReg, gsFormatoFecha)
End If

If Mid(psOpeCod, 3, 1) = "2" Then
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
                          ByVal psCtaIfCod As String, ByVal pnMovNro As Long, ByRef oMov As DMov)



SQL = " SELECT  DISTINCT M.NMOVNRO, M.CMOVNRO , MC.NMOVIMPORTE" _
    & " FROM    MOV M " _
    & "         JOIN MOVCTAIF MC ON MC.NMOVNRO = M.NMOVNRO " _
    & "         JOIN CTAIF C     ON MC.CPERSCOD=C.CPERSCOD AND MC.CIFTPO = C.CIFTPO AND MC.cCtaIFCod=C.cCtaIFCod " _
    & " Where   MC.nConceptoCod =" & CGCtaIFConInteres & " And MC.NMOVIMPORTE > 0 And m.nMovFlag = " & gMovFlagVigente & " " _
    & "         AND SUBSTRING(M.CMOVNRO,1,8)> CONVERT(CHAR(8),dCtaIfCap,112) " _
    & "         AND MC.cPersCod='" & psPersCod & "' AND MC.cIFTpo='" & Format(pnIFTpo, "00") & "' AND MC.cCtaIFCod ='" & psCtaIfCod & "'"

Set rs = oMov.CargaRecordSet(SQL)
Do While Not rs.EOF
    oMov.InsertaMovRef pnMovNro, rs!nMovNro
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
End Sub
Public Function GrabaCapitalizaIFEfectivo(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                         ByVal rsBillete As ADODB.Recordset, ByVal rsMoneda As ADODB.Recordset, _
                                         ByVal psCtaDebe As String, ByVal pnImporte As Currency, _
                                         ByVal psCtaDif As String, ByVal pnImporteDif As Currency, _
                                         ByVal psCtaHaberInt As String, ByVal pnImporteInt As Currency, _
                                         ByVal psCtaHaberIntCal As String, ByVal pnImporteIntCalc As Currency, _
                                         ByVal psCtaHaberCapital As String, ByVal pnImporteCapital As Currency, _
                                         ByVal psPersCod As String, ByVal psIFTpo As String, _
                                         ByVal psCtaIfCod As String, ByVal pdCtaIFCap As Date, Optional pnDiasTrans As Integer = 0) As Integer
                    
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
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDif, pnImporteDif
End If

If pnImporteInt > 0 Then
    lnMovItem = lnMovItem + 1:
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaberInt, pnImporteInt * -1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, psPersCod, psIFTpo, psCtaIfCod
End If
If pnImporteIntCalc > 0 Then
    lnMovItem = lnMovItem + 1:
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaberIntCal, pnImporteIntCalc * -1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, psPersCod, psIFTpo, psCtaIfCod
End If

'Referencia a todos los Movimientos de Intereses despues de la Ultima Capitalizacion
'InsertaRefCap psPersCod, psIFTpo, psCtaIfCod, lnMovNro, oDMov
If pnImporteCapital > 0 And psCtaHaberCapital <> "" Then
    lnMovItem = lnMovItem + 1:
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaberCapital, pnImporteCapital * -1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, 1, psPersCod, psIFTpo, psCtaIfCod
    
    'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
    'oDMov.InsertaMovCtaIF lnMovNro, psPersCod, psIFTpo, psCtaIfCod, CGCtaIFConCapital, pnImporteCapital * -1
    'oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                            psPersCod, psIFTpo, psCtaIfCod
End If
oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIfCod, , , Format(pdCtaIFCap, gsFormatoFecha), , , gEstadoCtaIFCancelada, 0, psMovNro

If Mid(psOpeCod, 3, 1) = "2" Then
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
Public Function GrabaCapitalizaDoc(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
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

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc
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

If pnImporteInt > 0 Then
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaberInt, pnImporteInt * -1
End If
If pnImporteIntCalc > 0 Then
    lnMovItem = lnMovItem + 1:
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaberIntCal, pnImporteIntCalc * -1
End If
''INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
'oDMov.InsertaMovCtaIF lnMovNro, psPersCod, psIFTpo, psCtaIfCod, CGCtaIFConInteres, (pnImporteInt + pnImporteIntCalc) * -1, pnDiasTrans

'InsertaRefCap psPersCod, psIFTpo, psCtaIfCod, lnMovNro, oDMov
oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIfCod, , , Format(pdCtaIFCap, gsFormatoFecha), , , , 0, psMovNro
If pnImporteCapital > 0 And psCtaHaberCapital <> "" Then
    lnMovItem = lnMovItem + 1:
    lnMovOrden = 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaberCapital, pnImporteCapital * -1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIfCod
    
    'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
    'oDMov.InsertaMovCtaIF lnMovNro, psPersCod, psIFTpo, psCtaIfCod, CGCtaIFConCapital, pnImporteCapital * -1
    'oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
                            psPersCod, psIFTpo, psCtaIfCod
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

If Mid(psOpeCod, 3, 1) = "2" Then
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
Public Function GrabaPagoSunat(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                              ByVal pnTpoDocPago As TpoDoc, ByVal psNroDocPago As String, ByVal psFechaDocPago As String, _
                              ByVal rsBillete As ADODB.Recordset, ByVal rsMoneda As ADODB.Recordset, _
                              ByVal rsDebe As ADODB.Recordset, _
                              ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                              ByVal psIFPersCod As String, ByVal psIFTpo As String, _
                              ByVal psCtaIfCod As String, _
                              ByVal pnTpoDocFormulario As TpoDoc, ByVal psNroDocFormulario As String, ByVal pdFechaDocFormulario As Date, _
                              ByVal psDocVoucher As String) As Integer
            
Dim lnMovItem  As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lsMsgErr As String
Set oDMov = New DMov
On Error GoTo GrabaPagoSunatErr

GrabaPagoSunat = 1

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable
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
      oDMov.ActualizaSaldoCtaIF CDate(psFechaDocPago), psIFPersCod, psIFTpo, psCtaIfCod
      oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
      oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psIFPersCod, psIFTpo, psCtaIfCod
      oDMov.InsertaMovDoc lnMovNro, pnTpoDocPago, psNroDocPago, psFechaDocPago
      oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocVoucher, psFechaDocPago
End Select

'Grabamos los datos del Formulario a Pagar
oDMov.InsertaMovDoc lnMovNro, pnTpoDocFormulario, psNroDocFormulario, Format(pdFechaDocFormulario, gsFormatoFecha)
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

Public Function GrabaCartaFianza(psMovNro As String, psOpeCod As String, psMovDesc As String, pnTpoDocCartaFianza As TpoDoc, psNumCarta As String, psFechaEmision As String, _
                                psCodBanco As String, psCodPers As String, psFechaVenc As String, psCtaCodDebe As String, psCtaCodHaber As String, pnImporte As Currency) As Integer
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lsMsgErr As String
Dim oDMov As DMov
Set oDMov = New DMov
On Error GoTo GrabaCartaFianzaErr

GrabaCartaFianza = 1

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
oDMov.InsertaMovDoc lnMovNro, pnTpoDocCartaFianza, psNumCarta, psFechaEmision
oDMov.InsertaMovCta lnMovNro, 1, psCtaCodDebe, pnImporte
oDMov.InsertaMovCta lnMovNro, 2, psCtaCodHaber, pnImporte * -1
'oDMov.InsertaMovOtrosItem lnMovNro, 1, "CF", pnImporte, psCodBanco
'oDMov.InsertaMovOtrosItem lnMovNro, 2, "CF", pnImporte, Format(CDate(psFechaVenc), gsFormatoFecha)
oDMov.InsertaMovCartaFianza lnMovNro, psCodPers, psCodBanco, Format(CDate(psFechaVenc), gsFormatoFecha)
oDMov.InsertaMovGasto lnMovNro, psCodPers, "0"

If Mid(psOpeCod, 3, 1) = "2" Then
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

Public Function GrabaCartaFianzaSalida(pdFecha As Date, psCodAge As String, psCodUser As String, psOpeCod As String, psMovDesc As String, rsCarta As ADODB.Recordset, psCtaCodDebe As String, psCtaCodHaber As String) As String
Dim lbTrans  As Boolean
Dim lsMovNro As String
Dim lnMovNro As Long
Dim lsMsgErr As String
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
            oDMov.InsertaMov lsMovNro, psOpeCod, psMovDesc
            lnMovNro = oDMov.GetnMovNro(lsMovNro)
            oDMov.InsertaMovCont lnMovNro, rsCarta!Monto, 0, "0"
            oDMov.InsertaMovCta lnMovNro, 1, psCtaCodDebe, rsCarta!Monto
            oDMov.InsertaMovCta lnMovNro, 2, psCtaCodHaber, rsCarta!Monto * -1
            oDMov.InsertaMovRef lnMovNro, rsCarta!nMovNro
            If Mid(psOpeCod, 3, 1) = "2" Then
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

Public Function GetDatosCartaFianza(psCtaContCod As String, Optional pnTipo As Integer = 0, Optional psFechaDel As String = "", Optional psFechaAl As String = "") As ADODB.Recordset
Dim lsFiltro As String

Select Case pnTipo
   Case 0   'Todas las Cartas Vigentes
      lsFiltro = " and Ref.nMovNro is NULL "
   Case 1   'Cartas Recibidas según Rango de Fechas
      lsFiltro = " and LEFT(M.cMovNro,8) BETWEEN '" & psFechaDel & "' and '" & psFechaAl & "' "
   Case 2   'Cartas Entregadas según Rango de Fechas
      lsFiltro = " and LEFT(Ref.cMovNro,8) BETWEEN '" & psFechaDel & "' and '" & psFechaAl & "' "
End Select
oConect.AbreConexion
SQL = " SELECT MD.CDOCNRO, CONVERT(CHAR(30),M.CMOVDESC) DESCRIPCION , Bco.cPersNombre cBanco, P.cPersNombre, " _
    & "        ISNULL(MME.NMOVMEIMPORTE, MC.NMOVIMPORTE) *-1  NMOVIMPORTE , MD.DDOCFECHA, CONVERT(datetime,MCF.dCartaVenc) AS FECHAVENC, MCF.cIFPersCod, M.cMovNro, M.nMovNro, ISNULL(Ref.cMovNro,'') cMovNroSalida " _
    & " FROM MOVCTA MC INNER JOIN MOV M ON M.nMOVNRO=MC.nMOVNRO " _
    & "    INNER JOIN MOVDOC       MD  ON MD.nMOVNRO=M.nMOVNRO " _
    & "    INNER JOIN MOVCARTAFIANZA MCF ON MCF.nMOVNRO=M.nMOVNRO " _
    & "    INNER JOIN Persona BCO ON BCO.CPERSCOD = MCF.cIFPersCod " _
    & "    INNER JOIN Persona P   ON P.cPersCod=MCF.cPersCod " _
    & "     LEFT JOIN (SELECT MR.nMOVNRO, MR.nMovNroRef, M1.cMovNro FROM MOVREF MR JOIN Mov M1 ON M1.nMovNro = mr.nMovNro " _
    & "           WHERE  M1.nMOVESTADO= " & gMovEstContabMovContable & " AND M1.nMOVFLAG NOT  IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
    & "          ) Ref ON Ref.nMovNroRef = m.nMovNro " _
    & "     LEFT JOIN MOVME MME ON MME.nMOVNRO=MC.nMOVNRO AND MC.nMOVITEM=MME.nMOVITEM " _
    & " WHERE MC.NMOVIMPORTE < 0 and MC.CCTACONTCOD = '" & Trim(psCtaContCod) & "' AND M.nMOVESTADO= " & gMovEstContabMovContable & " AND M.nMOVFLAG NOT  IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
    & lsFiltro
    
Set GetDatosCartaFianza = oConect.CargaRecordSet(SQL)
End Function


Public Function GrabaPagoProveedor(psMovNro As String, psOpeCod As String, psMovDesc As String, psCtaContDebeB As String, psCtaContDebeS As String, _
                                   psCtaContHaber As String, pnImporteB As Currency, pnImporteS As Currency, psPersCod As String, psTpoIf As String, psPersCodIF As String, psCtaIf As String, _
                                   prsBilletaje As ADODB.Recordset, psDocTpo As String, psDocNro As String, psFechaDoc As String, psDocVoucher As String, rsDocs As ADODB.Recordset, psCuentaAho As String, _
                                   ByVal psCtaDiferencia As String, ByVal pnMontoDiferencia As Currency) As Integer
Dim lnMovItem  As Long
Dim lnMovOrden As Integer
Dim lnMovNro As Long
Dim oDMov    As DMov
Dim lbTrans  As Boolean
Dim lsMovNro As String
Dim lsMsgErr As String
Set oDMov = New DMov
On Error GoTo GrabaPagoProveedorErr
oDMov.BeginTrans
lbTrans = True
GrabaPagoProveedor = 1

oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporteB + pnImporteS, "0", "0"

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

'guardamos el billetaje si este posee
If Not prsBilletaje Is Nothing Then
   oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, prsBilletaje, psCtaContHaber, "H"
Else
    'guardamos la cuenta del haber
    lnMovItem = lnMovItem + 1: lnMovOrden = 0
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber, (pnImporteB + pnImporteS) * -1
    If psPersCodIF <> "" Then
        'guardamos la Cta de la IF si es que utiliza esta
        lnMovOrden = lnMovOrden + 1
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCodIF, psTpoIf, psCtaIf
        
        'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
        'oDMov.InsertaMovCtaIF lnMovNro, psPersCodIf, psTpoIf, psCtaIf, CGCtaIFConCapital, (pnImporteB + pnImporteS) * -1
        oDMov.ActualizaSaldoCtaIF CDate(GetFechaMov(psMovNro, True)), psPersCodIF, psTpoIf, psCtaIf
    End If
End If
If psDocNro <> "" Then
   oDMov.InsertaMovDoc lnMovNro, psDocTpo, psDocNro, psFechaDoc
End If
If psDocVoucher <> "" Then
   oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocVoucher, GetFechaMov(psMovNro, False)
End If
If psDocTpo = TpoDocNotaAbono Or psDocTpo = TpoDocNotaCargo Then
    oDMov.InsertaRegDocCuenta psDocTpo, psDocNro, psCuentaAho
'    Dim oCon As New DConecta
'    If oCon.AbreConexionRemota(Left(psCuentaAho, 2)) Then
'        oCon.BeginTrans
'        lnSaldo = Abono(psCuentaAho, CCur(lsMonto), lsOpeAbonoCod, "112" & Left(psCuentaAho, 2), pgsCodUser, lsNotaAbono, "PAGO : " & psPlanillaDes & " - " & psPlanillaPeriodo, oCon, pgdFecSis)
'        lsCadBol = lsCadBol & ImprimeBoletaCad(pgdFecSis, "ABONO RRHH", "Depósito RRHH*Nro." & lsNotaAbono, "", lsMonto, prRs!Nombre, psCuentaAho, "", CCur(lnSaldo), 0, "Nota Abono", 0, 0, False, False, , , , , , , , False) & oImpresora.gPrnSaltoPagina
'        oCon.CierraConexion
'    End If
'    Set oCon = Nothing
End If
oDMov.InsertaMovGasto lnMovNro, psPersCod, "0"

If pnMontoDiferencia <> 0 Then
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDiferencia, pnMontoDiferencia * -1
End If

'Referenciamos todos los Documentos que se Pagan
If Not rsDocs Is Nothing Then
   Do While Not rsDocs.EOF
      If rsDocs!Itm = "1" Then
         oDMov.InsertaMovRef lnMovNro, rsDocs!nMovNro
      End If
      rsDocs.MoveNext
   Loop
End If
If Mid(psOpeCod, 3, 1) = 2 Then
    oDMov.GeneraMovME lnMovNro, psMovNro
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
    Set oDMov = Nothing
    Err.Raise vbObjectError + 100, "GrabaPagoProveedor", lsMsgErr
End Function

Public Function GrabaPagoProveedorEntregaDoc(psMovNro As String, psOpeCod As String, psMovDesc As String, prs As ADODB.Recordset) As Integer
Dim lnMovItem  As Integer
Dim lnMovNro As Long
Dim oDMov    As DMov
Dim lbTrans  As Boolean
Dim lsMovNro As String
Dim lsMsgErr As String
Set oDMov = New DMov
On Error GoTo GrabaPagoProveedorErr
oDMov.BeginTrans
lbTrans = True
GrabaPagoProveedorEntregaDoc = 1

oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
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
Dim SQL      As String
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
SQL = " SELECT  md.dDocFecha, d.cDocAbrev, md.nDocTpo, md.cDocNro, p.cPersNombre, m.cMovDesc, " _
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
    
Set GetDatosPagoProveedor = oConect.CargaRecordSet(SQL)

Exit Function
GetDatosPagoProveedorErr:
    Err.Raise vbObjectError + 100, "GetDatosPagoProveedor", Err.Description
End Function


'Public Function GetOrdPagFondoFijoEntregado(ByVal nMoneda As Moneda, ByVal sNumOrdPag As String) As adodb.Recordset
'Dim oCaj As DCajaGeneral
'Set oCaj = New DCajaGeneral
'Set GetOrdPagFondoFijoEntregado = oCaj.GetOrdPagFondoFijoEntregado(nMoneda, sNumOrdPag)
'Set oCaj = Nothing
'End Function

Public Function GetOrdPagFondoFijoEntregado(ByVal nMoneda As Moneda, ByVal sNumOrdPag As String) As ADODB.Recordset
Dim sSql As String
Dim rstemp As ADODB.Recordset

Set rstemp = New ADODB.Recordset

sSql = "SELECT A.nMovNro,B.dDocFecha, G.cDocAbrev, B.nDocTpo cDocTpo, B.cDocNro, E.cPersNombre cNomPers, A.cMovDesc, D.cPersCod cObjetoCod, " _
    & "A.cMovNro, ISNULL(ME.nMovMEImporte, C.nMovImporte) * -1 As nDocImporte " _
    & "FROM Mov A JOIN MovDoc B ON B.nMovNro = A.nMovNro " _
    & "     JOIN MovCta C ON c.nMovNro = a.nMovNro LEFT JOIN MovMe ME ON ME.nMovNro = C.nMovNro And ME.nMovItem = C.nMovItem " _
    & "     JOIN MovGasto D ON D.nMovNro = A.nMovNro " _
    & "     JOIN Persona E ON E.cPersCod = D.cPersCod " _
    & "     JOIN Documento G ON G.nDocTpo = B.nDocTpo " _
    & "WHERE A.nMovEstado = " & gMovEstContabMovContable & "  and A.nMovFlag <> " & gMovFlagEliminado & "  And B.nDoctpo = " & TpoDocOrdenPago & " "
    
    'A rendir
    sSql = sSql & " And (C.cCtaContCod LIKE '21_10502%' or c.cctacontcod like '29_80701%') and (copecod like  '40_132' OR copecod like  '40_322' or  copecod like  '42_113' or  copecod like  '40_232' or  copecod like  '42_009')   "
    sSql = sSql & " and c.cCtaContCod like '__" & nMoneda & "%' "
                '& "And C.cCtaContCod LIKE '21_10502%' and c.cCtaContCod like '__" & nMoneda & "%' "
    'Fin a rendir

sSql = sSql & "And C.nMovImporte < 0 And B.cDocNro = '" & sNumOrdPag & "'" _
    & "And EXISTS (SELECT H.nMovNro FROM MovRef H WHERE H.nMovNro = A.nMovNro) " _
    & "And NOT EXISTS (SELECT MP.nMovNro FROM MovRef Pag JOIN Mov MP on MP.nMovNro = Pag.nMovNro " _
    & "    WHERE Pag.nMovNroRef = A.nMovNro And MP.nMovFlag <> " & gMovFlagEliminado & "  And MP.cOpeCod IN ('" & gAhoRetFondoFijo & "','" & gAhoRetFondoFijoCanje & " ')) " _
    & "ORDER BY a.cMovNro, e.cPersNombre "
    
    rstemp.CursorLocation = adUseClient
    rstemp.Open sSql, oConect.ConexionActiva, adOpenStatic, adLockOptimistic, adCmdText
    
    
Set GetOrdPagFondoFijoEntregado = rstemp
   Set rstemp = Nothing

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

Set oConect = New DConecta
oConect.AbreConexion
End Sub

Private Sub Class_Terminate()
oConect.CierraConexion
Set oConect = Nothing
End Sub

Public Function GetCajeroDestino(pnMovNroHabDev As Long) As ADODB.Recordset
    Dim SQL As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    SQL = " Select Top 1 RH.cUser, dbo.PstaNombre(PE.cPersNombre,1) Nombre From MovHabDevCajero MHDC" _
        & " Inner Join RRHH RH On RH.cUser = MHDC.cUsuDest" _
        & " Inner Join Persona PE On RH.cPersCod = PE.cPersCod" _
        & " Where MHDC.nMovNro = " & pnMovNroHabDev
    
    Set GetCajeroDestino = oCon.CargaRecordSet(SQL)
End Function



