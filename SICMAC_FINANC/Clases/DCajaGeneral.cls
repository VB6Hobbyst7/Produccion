VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCajaGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsServerCom As String
Dim vsServerPers As String
Dim vsConexion As String
Dim oConect As DConecta
Dim pnBitCentral As Boolean

Public Function InsertaOpCajaGen(ByVal psRangoCod As String, ByVal pdFechaIng As Date, ByVal pnRangoEmiIni As Long, ByVal pnRangoEmiFin As Long, ByVal psEstadoRango As CGEstadosOp, ByVal psUltActualizacion As String, ByVal pnMoneda As Moneda, Optional pbEjectBacth As Boolean = False) As Integer
Dim sql As String
On Error GoTo InsertaOpCajaGenErr

InsertaOpCajaGen = 1
sql = "INSERT INTO OPCAJAGEN(cRangoCod, dFechaIng, nRangoEmiIni, nRangoEmiFin, cEstadoRango, cUltActualizacion, cMoneda) " _
    & " VALUES ('" & psRangoCod & "','" & Format(pdFechaIng, gsFormatoFecha) & "'," & pnRangoEmiIni & "," & pnRangoEmiFin & ",'" & psEstadoRango & "','" & psUltActualizacion & "','" & pnMoneda & "')"

If pbEjectBacth Then
    oConect.AdicionaCmdBatch sql
Else
    oConect.Ejecutar sql
End If
InsertaOpCajaGen = 0
Exit Function
InsertaOpCajaGenErr:
    Call RaiseError(MyUnhandledError, "DCajaGeneral:InsertaOpCajaGen")

End Function
Public Function InsertaOpDetCajaGen(ByVal psRangoCod As String, ByVal pnRangoIni As Long, ByVal pnRangoFin As Long, ByVal psEstado As CGEstadosOp, ByVal psDescripcion As String, ByVal psUltActualizacion As String, Optional pbEjectBacth As Boolean = False) As Integer
Dim sql As String
On Error GoTo InsertaOpDetCajaGenErr

InsertaOpDetCajaGen = 1
sql = "INSERT INTO OPCAJAGENDET(cRangoCod,nRangoIni,nRangoFin,cEstadoDet,cDescDet,cUltActualizacion) " _
    & " VALUES ('" & psRangoCod & "'," & pnRangoIni & "," & pnRangoFin & ",'" & psEstado & "','" & Replace(psDescripcion, "'", "''") & "','" & psUltActualizacion & "')"

If pbEjectBacth Then
    oConect.AdicionaCmdBatch sql
Else
    oConect.Ejecutar sql
End If
InsertaOpDetCajaGen = 0
Exit Function
InsertaOpDetCajaGenErr:
    Call RaiseError(MyUnhandledError, "DCajaGeneral:InsertaOpDetCajaGen")
End Function
Public Function EliminaOpCajaGen(ByVal psRangoCod As String, Optional pbEjectBacth As Boolean = False) As Integer
Dim sql As String
On Error GoTo EliminaOpCajaGenErr
sql = "DELETE OPCAJAGEN WHERE cRangoCod= '" & psRangoCod & "'"
EliminaOpCajaGen = 1
If pbEjectBacth Then
    oConect.AdicionaCmdBatch sql
Else
    oConect.Ejecutar sql
End If
EliminaOpCajaGen = 0

Exit Function
EliminaOpCajaGenErr:
    Call RaiseError(MyUnhandledError, "DCajaGeneral:EliminaOpCajaGen")

End Function
Public Function EliminaOpDetCajaGen(ByVal psRangoCod As String, ByVal pnRangoIni As Long, ByVal pnRangoFin As Long, Optional pbEjectBacth As Boolean = False) As Integer
Dim sql As String
On Error GoTo EliminaOpDetCajaGenErr

EliminaOpDetCajaGen = 1
sql = "DELETE FROM OPCAJAGENDET WHERE cRangoCod='" & psRangoCod & "' AND nRangoIni=" & pnRangoIni & " AND nRangoFin=" & pnRangoFin

If pbEjectBacth Then
    oConect.AdicionaCmdBatch sql
Else
    oConect.Ejecutar sql
End If
EliminaOpDetCajaGen = 0


Exit Function
EliminaOpDetCajaGenErr:
    Call RaiseError(MyUnhandledError, "DCajaGeneral:EliminaOpDetCajaGen ")

End Function
Public Function EjecutaBatch() As Integer
    EjecutaBatch = oConect.EjecutarBatch
End Function
Private Sub Class_Initialize()
    Dim oConst As NConstSistemas
    Set oConst = New NConstSistemas
    pnBitCentral = IIf(oConst.LeeConstSistema(gConstSistBitCentral) = "1", True, False)
    
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
If oConect.AbreConexion = False Then
    Call RaiseError(MyUnhandledError, "DCajaGeneral:Initialize Method. Error en Conexion a Base de datos")
End If
End Sub
Private Sub Class_Terminate()
oConect.CierraConexion
Set oConect = Nothing
End Sub
Public Function ActualizaOpCajaGen(ByVal psRangoCod As String, ByVal pnRangoEmiIni As Long, ByVal pnRangoEmiFin As Long, ByVal psUltActualizacion As String, Optional pbEjectBacth As Boolean = False) As Integer
Dim sql As String
On Error GoTo ActualizaOpCajaGenErr

ActualizaOpCajaGen = 1
sql = "UPDATE OPCajaGen SET nRangoEmiIni=" & pnRangoEmiIni & ",nRangoEmiFin=" & pnRangoEmiFin & ", cUltActualizacion='" & psUltActualizacion & "' WHERE cRangoCod='" & psRangoCod & "'"
If pbEjectBacth Then
    oConect.AdicionaCmdBatch sql
Else
    oConect.Ejecutar sql
End If
ActualizaOpCajaGen = 0
Exit Function
ActualizaOpCajaGenErr:
    Call RaiseError(MyUnhandledError, "DCajaGeneral:ActualizaOpCajaGen ")
End Function

'Public Function GetDatosProvisionesProveedores(psCtasCod As String, psDocs As String, pdFechaIni As Date, pdFechaFin As Date, Optional pbPend As Boolean = False, Optional bIncluirMonto As Boolean = False, Optional pnMonto As Double = 0) As ADODB.Recordset
'Dim sql As String
'On Error GoTo GetDatosProvisionesProveedoresErr
'sql = "SELECT distinct md.dDocFecha, doc.cDocAbrev, md.nDocTpo, md.cDocNro, Prov.cPersNombre cPersona, Prov.nPersPersoneria, M.cMovDesc, Prov.cPersCod, " _
'     & "       m.cMovNro, m.nMovNro, mc.cCtaContCod, ISNULL(mrend.nSaldo, ISNULL(me.nMovMeImporte,mc.nMovImporte) * -1) as nMovImporte, mc.nMovImporte * -1 as nMovImporteSoles " _
'     & "FROM   Mov m JOIN MovDoc md ON md.nMovNro = m.nMovNro " _
'     & "             JOIN MovCta mc ON mc.nMovNro = m.nMovNro LEFT JOIN MovMe ME ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'     & "             JOIN MovGasto mg ON mg.nMovNro = m.nMovNro " _
'     & "        LEFT JOIN (SELECT mr.nMovNro, mr.nMovNroRef FROM MovRef mr JOIN Mov m1 ON m1.nMovNro = mr.nMovNro " _
'     & "                   WHERE m1.nMovEstado = " & gMovEstContabMovContable & " and m1.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') and RTRIM(ISNULL(mr.cAgeCodRef,'')) = '' " _
'     & "                  and m1.cOpeCod not in (" & gOpeCGOpeBancosRetCtasBancosDetracMN & "," & gOpeCGOpeBancosRetCtasBancosDetracME & "," & gOpeCGOpeBancosRetCtasBancosEmbargoMN & ", " & gOpeCGOpeBancosRetCtasBancosEmbargoME & ")) ref ON  ref.nMovNroRef = m.nMovNro " _
'     & "             JOIN Persona Prov  ON Prov.cPersCod = mg.cPersCod " _
'     & "             JOIN Documento Doc ON Doc.nDocTpo = md.nDocTpo " _
'     & "        LEFT JOIN MovPendientesRend mrend ON mrend.nMovNro = m.nMovNro and mrend.cCtaContCod IN (" & psCtasCod & ") " _
'     & "WHERE  m.nMovEstado IN (" & gMovEstContabMovContable & "," & gMovEstContabPendiente & ") and m.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
'     & "       and (ref.nMovNro IS NULL or ISNULL(mrend.nSaldo,0) <> 0) and " _
'     & " md.dDocFecha BETWEEN '" & Format(pdFechaIni, gsFormatoFecha) & "' AND '" & Format(pdFechaFin, gsFormatoFecha) & "' " _
'     & "       and md.nDocTpo IN (" & psDocs & ") and mc.cCtaContCod IN (" & psCtasCod & ") and (mc.nMovImporte < 0 or (mc.nMovImporte > 0 And md.nDocTpo in (7)) ) "
'
'     If bIncluirMonto = True Then
'        If Mid(psCtasCod, 4, 1) = Moneda.gMonedaNacional Then
'            sql = sql & " And abs(mc.nMovImporte)=" & pnMonto & " "
'        Else
'            sql = sql & " And abs(isnull(me.nMovMeImporte,0))=" & pnMonto & " "
'        End If
'     End If
'
'sql = sql & "ORDER BY md.dDocFecha,Prov.cPersNombre, m.nMovNro"
'Set GetDatosProvisionesProveedores = oConect.CargaRecordSet(sql)
'Exit Function
'GetDatosProvisionesProveedoresErr:
'    Call RaiseError(MyUnhandledError, "DCajaGeneral:GetDatosProvisionesProveedores")
'End Function
'EJVG20131119 Se agregó pdFechaEmisionIni

Public Function GetProveedoresCtasAperturadasSBS(psOpecod As String, Optional psOpt As String) As String
Dim sql As String
Dim oConec As New DConecta
Dim oRs As New ADODB.Recordset
Dim lsCtaProvCtaDet As String
On Error GoTo ErroGetProveedoresCtasAperturadasSBS
sql = "Exec stp_sel_ObtieneProveedoresCtasAperturadasSBS '" & psOpecod & "', '" & psOpt & "'"
oConec.AbreConexion
Set oRs = oConec.CargaRecordSet(sql)
If Not oRs.BOF And Not oRs.EOF Then
    lsCtaProvCtaDet = oRs!cCtaComp
Else
    lsCtaProvCtaDet = ""
End If
oConec.CierraConexion
GetProveedoresCtasAperturadasSBS = lsCtaProvCtaDet
Exit Function
ErroGetProveedoresCtasAperturadasSBS:
    Call RaiseError(MyUnhandledError, "Carga ProveedoresCtasAperturadasSBS")
End Function '***NAGL INC1712260008 20171227

Public Function GetProveedoresCtasAperturadasSBSbyItem(psOpecod As String, Optional psOpt As String) As ADODB.Recordset
Dim sql As String
Dim oConec As New DConecta
On Error GoTo ErroGetProveedoresCtasAperturadasSBSbyItem
oConec.AbreConexion
sql = "Exec stp_sel_ObtieneProveedoresCtasAperturadasSBS '" & psOpecod & "', '" & psOpt & "'"
Set GetProveedoresCtasAperturadasSBSbyItem = oConec.CargaRecordSet(sql)
oConec.CierraConexion
Exit Function
ErroGetProveedoresCtasAperturadasSBSbyItem:
    Call RaiseError(MyUnhandledError, "Carga ProveedoresCtasAperturadasSBS by Item")
End Function '***NAGL INC1712260008 20171227

'ALPA 20090318********************************************************************************************************************************************************
'Public Function GetDatosProvisionesProveedores(psCtasCod As String, psDocs As String, pdFechaIni As Date, pdFechaFin As Date, Optional pbPend As Boolean = False, Optional bConsiderarDetra As Integer = 0, Optional cCtaDetraccion As String = "29180799", Optional psCtasCodME As String = "", Optional psTipoInterfaz As String = "", Optional psArchivo As String = "", Optional psTipo As String = "", Optional pbPagoAgencia As Boolean = False, Optional psAgeCod As String, Optional psTpoB As String = "T") As ADODB.Recordset
Public Function GetDatosProvisionesProveedores(psCtasCod As String, psDocs As String, pdFechaIni As Date, pdFechaFin As Date, Optional pbPend As Boolean = False, Optional bConsiderarDetra As Integer = 0, Optional cCtaDetraccion As String = "29180799", Optional psCtasCodME As String = "", Optional psTipoInterfaz As String = "", Optional psArchivo As String = "", Optional psTipo As String = "", Optional pbPagoAgencia As Boolean = False, Optional psAgeCod As String, Optional psTpoB As String = "T", Optional pnTipoPago As Integer, Optional ByVal pbPagoNuevo As Boolean = False, Optional ByVal pdFechaEmisionIni As Date = "01/01/1900", Optional ByVal psListaIFiCod As String = "") As ADODB.Recordset
'**********************************************************************************************************************************************************************
Dim sql         As String
Dim sqltable    As String
Dim nTieneDetra As New DMov
Dim sTabla1     As String
Dim sTabla2     As String
Dim sTabla3     As String
Dim sTabla4     As String
'EJVG20131119 ***
Dim ldFechaEmisionFin As Date
If pbPagoNuevo Then
    ldFechaEmisionFin = DateAdd("m", 1, pdFechaEmisionIni)
End If
'END EJVG *******
On Error GoTo GetDatosProvisionesProveedoresErr
    
    psCtasCod = psCtasCod & IIf(psCtasCodME = "", "", "," & psCtasCodME)
    'sTabla1 = "##Tmp2" & Format(FechaHora(pdFechaFin), "yyyymmddhhmmss") & gsCodUser
    sTabla2 = "##Tmp1" & Format(FechaHora(pdFechaFin), "yyyymmddhhmmss") & gsCodUser
   ' sTabla3 = "##Tmp3" & Format(FechaHora(pdFechaFin), "yyyymmddhhmmss") & gsCodUser
    sTabla4 = "##Tmp4" & Format(FechaHora(pdFechaFin), "yyyymmddhhmmss") & gsCodUser
'Se comento porque demoraba mucho GITU
    'Primera Tabla
'    sql = " Select C.nMovOrigen "
'    sql = sql & "   Into " & sTabla1 & " From"
'    sql = sql & "     ( Select C1.nMovNro as nMovOrigen, C2.nMovNro AS nMovDestino FROM"
'        sql = sql & " ( select distinct mK.nMovNro "
'        sql = sql & " from MovCta MK Inner Join Mov MV on MK.nmovnro=Mv.nmovnro "
'        sql = sql & " where MK.cCtaContCod Like '" & cCtaDetraccion & "' and MK.nMovImporte <0 and MV.nmovflag=0 "
'        sql = sql & " ) C1 "
'    sql = sql & "       Left Join"
'        sql = sql & " (Select MR.nMovNro, MR.nMovNroRef "
'        sql = sql & " From MovRef MR Inner Join Mov MV1 on MR.nmovnroref = MV1.nmovnro And IsNull(MR.cAgeCodRef,'') = '' ANd MV1.nMovFlag = 0 ANd MV1.nMovEstado = 10  "
'        sql = sql & " Inner Join Mov MV2 on MR.nmovnro = mv2.nMovNro And mv2.nMovFlag = 0 And mv2.nMovEstado = 10  "
'        sql = sql & "  ) C2 "
'    sql = sql & "       On C1.nMovNro=C2.nMovNroRef   where  C2.nMovNroRef is not null ) C"
'    sql = sql & "   Group By C.nMovOrigen having Count(C.nMovDestino)<>1"
'    sql = sql & "   Union "
'    sql = sql & "   Select Distinct MR1.nMovNroref"
'    sql = sql & "   From MovRef MR1 "
'    sql = sql & " Inner Join Mov MV1 on MR1.nmovnroref=MV1.nmovnro And IsNull(MR1.cAgeCodRef,'') = '' ANd MV1.nMovFlag = 0 ANd MV1.nMovEstado = 10 and mv1.copecod not in ('" & OpeCGOpeProvPagoListSUNAT & "') " ''" & OpeRegPenalidadMN & "','" & OpeRegPenalidadME & "',
'    sql = sql & " Inner Join Mov MV2 on MR1.nmovnro = mv2.nMovNro And mv2.nMovFlag = 0 And mv2.nMovEstado = 10  and mv2.copecod not in ('" & OpeCGOpeProvPagoPagUNAT & "') " ''" & OpeRegPenalidadMN & "','" & OpeRegPenalidadME & "',
'    sql = sql & "   Inner Join MovDoc MD on MR1.nMovNroRef=MD.nMovNro"
'    sql = sql & "   where Len(MR1.nMovNro) > 0 "
'    sql = sql & " and MV1.nMovFlag = 0 And mv2.nMovFlag = 0 "
'    sql = sql & " and MR1.nMovNroref not in ("
'    sql = sql & "   select distinct movcta.nMovNro"
'    sql = sql & "   From MovCta  Inner Join Mov On Movcta.nMovNro= Mov.nMovNro "
'    sql = sql & "   where cCtaContCod = '29180799' and nMovImporte < 0 and Mov.nMovFlag=0 "
'    sql = sql & "   )"
    
    sqltable = " Select C.nMovOrigen From "
    sqltable = sqltable & "     ( Select C1.nMovNro as nMovOrigen, C2.nMovNro AS nMovDestino FROM"
        'ALPA 20090318*******
        'sqltable = sqltable & " ( select distinct mK.nMovNro,MK.nTipoPago "
        sqltable = sqltable & " ( select distinct mK.nMovNro,isnull(MK.nTipoPago,0) nTipoPago "
        '********************
        sqltable = sqltable & " from MovCta MK Inner Join Mov MV on MK.nmovnro=Mv.nmovnro "
        sqltable = sqltable & " where MK.cCtaContCod Like '" & cCtaDetraccion & "' and MK.nMovImporte <0 and MV.nmovflag=0 "
        sqltable = sqltable & " ) C1 "
    sqltable = sqltable & "       Left Join"
        'ALPA 20090318****
        'sqltable = sqltable & " (Select MR.nMovNro, MR.nMovNroRef,MR.nTipoPago "
        sqltable = sqltable & " (Select MR.nMovNro, MR.nMovNroRef,isnull(MR.nTipoPago,0) nTipoPago "
        '****************
        sqltable = sqltable & " From MovRef MR Inner Join Mov MV1 on MR.nmovnroref = MV1.nmovnro And IsNull(MR.cAgeCodRef,'') = '' ANd MV1.nMovFlag = 0 ANd MV1.nMovEstado = 10  "
        sqltable = sqltable & " Inner Join Mov MV2 on MR.nmovnro = mv2.nMovNro And mv2.nMovFlag = 0 And mv2.nMovEstado = 10  "
        sqltable = sqltable & "  ) C2 "
    'ALPA 20090318******
    'sqltable = sqltable & "       On C1.nMovNro=C2.nMovNroRef   where  C2.nMovNroRef is not null ) C"
    sqltable = sqltable & "       On C1.nMovNro=C2.nMovNroRef   where  C2.nMovNroRef is not null and isnull(C1.nTipoPago,0)=isnull(C2.nTipoPago,0)) C"
    '*******************
    sqltable = sqltable & "   Group By C.nMovOrigen having Count(C.nMovDestino)<>1"
    sqltable = sqltable & "   Union "
    sqltable = sqltable & "   Select Distinct MR1.nMovNroref"
    sqltable = sqltable & "   From MovRef MR1 "
    sqltable = sqltable & " Inner Join Mov MV1 on MR1.nmovnroref=MV1.nmovnro And IsNull(MR1.cAgeCodRef,'') = '' ANd MV1.nMovFlag = 0 ANd MV1.nMovEstado = 10 and mv1.copecod not in ('" & OpeCGOpeProvPagoListSUNAT & "') " ''" & OpeRegPenalidadMN & "','" & OpeRegPenalidadME & "',
    sqltable = sqltable & " Inner Join Mov MV2 on MR1.nmovnro = mv2.nMovNro And mv2.nMovFlag = 0 And mv2.nMovEstado = 10  and mv2.copecod not in ('" & OpeCGOpeProvPagoPagUNAT & "') " ''" & OpeRegPenalidadMN & "','" & OpeRegPenalidadME & "',
    sqltable = sqltable & "   Inner Join MovDoc MD on MR1.nMovNroRef=MD.nMovNro"
    sqltable = sqltable & "   where Len(MR1.nMovNro) > 0 "
    sqltable = sqltable & " and MV1.nMovFlag = 0 And mv2.nMovFlag = 0 "
    sqltable = sqltable & " and MR1.nMovNroref not in ("
    sqltable = sqltable & "   select distinct movcta.nMovNro"
    sqltable = sqltable & "   From MovCta  Inner Join Mov On Movcta.nMovNro= Mov.nMovNro "
    'ALPA 20090318*****
    'sqltable = sqltable & "   where cCtaContCod = '29180799' and nMovImporte < 0 and Mov.nMovFlag=0 "
    sqltable = sqltable & "   where cCtaContCod = '29180799' and nMovImporte < 0 and Mov.nMovFlag=0 and isnull(nTipoPago,0)=isnull(MR1.nTipoPago,0) "
    '******************
    'ALPA 20090318*****
    'sqltable = sqltable & "   ) "
    sqltable = sqltable & "   ) and isnull(MR1.nTipoPago,0)=0 "
    '******************
    'oConect.Ejecutar sql
    'sql = " create index idxaaa on " & sTabla1 & " (nMovOrigen )"
    'oConect.Ejecutar sql
          
    'Segunda Tabla
    sql = "select M.nMovNro into " & sTabla2 & " From Mov M Inner Join MovDetra MD On M.nMovNro=MD.nMovNro Where M.nMovFlag=0 and MD.nEstado=2"
    oConect.Ejecutar sql
    sql = " create index idx_aa2 on " & sTabla2 & " (nMovNro)"
    oConect.Ejecutar sql
    

    'Tercera Tabla
'    sql = " SELECT mr.nMovNro, mr.nMovNroRef into " & sTabla3 & "  FROM MovRef mr JOIN Mov m1 ON m1.nMovNro = mr.nMovNro "
'    sql = sql & " WHERE m1.nMovEstado = " & gMovEstContabMovContable & " and m1.nMovFlag =0 and RTRIM(ISNULL(mr.cAgeCodRef,'')) = '' "
'    sql = sql & " and m1.cOpeCod not in ('" & gOpeCGOpeBancosRetCtasBancosDetracMN & "','" & gOpeCGOpeBancosRetCtasBancosDetracME & "','" & gOpeCGOpeBancosRetCtasBancosEmbargoMN & "','" & gOpeCGOpeBancosRetCtasBancosEmbargoME & "','" & OpeCGOpeProvPagoPagUNAT & "' ) "
    'sql = " SELECT mr.nMovNro, mr.nMovNroRef into " & sTabla3 & "  FROM VwnMovNroRef mr"
'    oConect.Ejecutar sql
'    sql = " create index idx_aa6 on " & sTabla3 & " (nMovNro)"
'    oConect.Ejecutar sql
'    sql = " create index idx_aa7 on " & sTabla3 & " (nMovNroRef)"
'    oConect.Ejecutar sql
    

    
    'Cuarta Tabla
    sql = " Select nDocTpo into " & sTabla4 & " from OpeDoc where cOpeCod = '421118'"
    oConect.Ejecutar sql
    sql = " create index idx_bb1 on " & sTabla4 & " (nDocTpo)"
    oConect.Ejecutar sql
    

        'ALPA 20090317******************************************************
        'sql = "SELECT distinct md.dDocFecha, doc.cDocAbrev, md.nDocTpo, md.cDocNro, Prov.cpersnombre cPersona, Prov.nPersPersoneria, M.cMovDesc, Prov.cPersCod, "
        sql = "SELECT distinct isnull(mc.nTipoPago,0) nPenalidad,md.dDocFecha, doc.cDocAbrev, md.nDocTpo, md.cDocNro, Prov.cpersnombre cPersona, Prov.nPersPersoneria, M.cMovDesc, Prov.cPersCod, "
        '*******************************************************************
        sql = sql & "       m.cMovNro, m.nMovNro, mc.cCtaContCod, ISNULL(me.nMovMeImporte ,mc.nMovImporte) * -1 as nMovImporte, (mc.nMovImporte )* -1 as nMovImporteSoles,Provi.cPersIDnro  " '+dbo.getMontoPenalidad(m.nMovNro) '+dbo.getMontoPenalidad(m.nMovNro)
        If psTipoInterfaz = "PAGO" Or psTipoInterfaz = "lbReporte" Then
            'Se comento la linea para que se puede pagar indistinatmente
            'los proveedores con pago sunat, y se agrego la linea siguiente. GITU
            'sql = sql & "   ,isnull(ps.nimportecoactivo,0) nImporteCoactivo, case when me.nMovMeImporte is NULL then (mc.nMovImporte*-1) else Round((me.nMovMeImporte*-1),2) end MontoPago, ps.nTpoCambio ,isnull(dbo.GetMontoPagadoSUNAT(m.nmovnro,''),0) MontoPagadoSUNAT,isnull(dbo.GetMontoPagadoSUNAT(m.nmovnro,'1'),0) MontoPagadoSUNATS ,'' MovEnvio , '' Agencia "
            sql = sql & "   ,isnull(ps.nimportecoactivo,0) nImporteCoactivo, case when me.nMovMeImporte is NULL then (mc.nMovImporte*-1-ps.nimportecoactivo) else Round((me.nMovMeImporte*-1)-(ps.nimportecoactivo/ps.ntpocambio),2) end MontoPago, ps.nTpoCambio ,isnull(dbo.GetMontoPagadoSUNAT(m.nmovnro,''),0) MontoPagadoSUNAT,isnull(dbo.GetMontoPagadoSUNAT(m.nmovnro,'1'),0) MontoPagadoSUNATS ,'' MovEnvio , '' Agencia "
            'EJVG20131120 ***
            If pbPagoNuevo Then
                sql = sql & " ,cPersCodIFi = ISNULL(CASE WHEN SUBSTRING(mc.cCtaContCod,3,1)='1' THEN C.cPersCodBancoMN ELSE C.cPersCodBancoME END,''),cPersNombreIFi = ISNULL(CASE WHEN SUBSTRING(mc.cCtaContCod,3,1)='1' THEN C.cPersNombreBancoMN ELSE C.cPersNombreBancoME END,''),cCtaCodIFi = ISNULL(CASE WHEN SUBSTRING(mc.cCtaContCod,3,1)='1' THEN C.cCtaCodMN ELSE C.cCtaCodME END,'') "
            End If
            'END EJVG *******
        End If
        sql = sql & " FROM   Mov m JOIN MovDoc md ON md.nMovNro = m.nMovNro "
        sql = sql & "             JOIN MovCta mc ON mc.nMovNro = m.nMovNro LEFT JOIN MovMe ME ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem "
        sql = sql & "             JOIN MovGasto mg ON mg.nMovNro = m.nMovNro "
        'sql = sql & "       LEFT JOIN (Select * from  " & sTabla3 & "  ) ref ON  ref.nMovNroRef = m.nMovNro "
        'sql = sql & "             LEFT JOIN (SELECT mr.nMovNro, mr.nMovNroRef FROM MovRef mr JOIN Mov m1 ON m1.nMovNro = mr.nMovNro "
        sql = sql & "             LEFT JOIN (SELECT mr.nMovNro, mr.nMovNroRef,isnull(mr.nTipoPago,0) nTipoPago FROM MovRef mr JOIN Mov m1 ON m1.nMovNro = mr.nMovNro "
        sql = sql & "                        WHERE m1.nMovEstado = " & gMovEstContabMovContable & " and m1.nMovFlag =0 and RTRIM(ISNULL(mr.cAgeCodRef,'')) = '' "
        sql = sql & "                               and m1.cOpeCod not in ('" & gOpeCGOpeBancosRetCtasBancosDetracMN & "','" & gOpeCGOpeBancosRetCtasBancosDetracME & "','" & gOpeCGOpeBancosRetCtasBancosEmbargoMN & "','" & gOpeCGOpeBancosRetCtasBancosEmbargoME & "','" & OpeCGOpeProvPagoPagUNAT & "' ) "
        'ALPA 20090318**************************************************************
        'sql = sql & "               ) ref ON  ref.nMovNroRef = m.nMovNro "
        sql = sql & "               ) ref ON  ref.nMovNroRef = m.nMovNro and isnull(ref.nTipoPago,0)= " & pnTipoPago & " "
        '***************************************************************************
        sql = sql & "             JOIN Persona Prov  ON Prov.cPersCod = mg.cPersCod "
        sql = sql & "             left join Persid ProvI  ON ProvI.cPersCod = mg.cPersCod and cPersIdTpo In (2) "
        sql = sql & "             JOIN Documento Doc ON Doc.nDocTpo = md.nDocTpo "
        sql = sql & "  Join " & sTabla4 & " OD On OD.nDocTpo = md.nDocTpo"

        If psTipoInterfaz = "PAGO" Then
            sql = sql & "          join MovControlPagoSunat ps ON ps.nmovnro = m.nMovNro and ps.cmovnrores is not null and ps.bvigente=1 and bvalido=1"
            'EJVG20131119 ***
            If pbPagoNuevo Then
                sql = sql & " Inner Join "
                sql = sql & " ("
                sql = sql & "   SELECT  M.nMovNro,P.cPersCod,P.cPersNombre,"
                sql = sql & "           cPersCodBancoMN = ISNULL(Prov.cPersCodBancoMN,''),cPersNombreBancoMN = ISNULL(PMN.cPersNombre,''),cctaCodMN = ISNULL(Prov.cctaCodMN,''),"
                sql = sql & "           cPersCodBancoME = ISNULL(Prov.cPersCodBancoME,''),cPersNombreBancoME = ISNULL(PME.cPersNombre,''),cCtaCodME = ISNULL(Prov.cCtaCodME,'')"
                sql = sql & "   FROM Mov M"
                sql = sql & "   LEFT JOIN MovRef MR ON MR.nMovNro = M.nMovNro"
                sql = sql & "   LEFT JOIN MovComprobante MComp ON MComp.nMovNro = MR.nMovNroRef"
                sql = sql & "   INNER JOIN MovGasto MG ON MG.nMovNro = M.nMovNro"
                sql = sql & "   INNER JOIN Persona P ON P.cPersCod = MG.cPersCod"
                sql = sql & "   LEFT JOIN Proveedor Prov ON Prov.cPersCod = P.cPersCod"
                sql = sql & "   LEFT JOIN Persona PMN ON PMN.cPersCod = Prov.cPersCodBancoMN"
                sql = sql & "   LEFT JOIN Persona PME ON PME.cPersCod = Prov.cPersCodBancoME"
                'sql = sql & "   union all" 'PASIERS0772014
'                sql = sql & "   SELECT  M.nMovNro,P.cPersCod,P.cPersNombre,"
'                sql = sql & "           cPersCodBancoMN = ISNULL(Prov.cPersCodBancoMN,''),cPersNombreBancoMN = ISNULL(PMN.cPersNombre,''),cctaCodMN = ISNULL(Prov.cctaCodMN,''),"
'                sql = sql & "           cPersCodBancoME = ISNULL(Prov.cPersCodBancoME,''),cPersNombreBancoME = ISNULL(PME.cPersNombre,''),cCtaCodME = ISNULL(Prov.cCtaCodME,'')"
'                sql = sql & "   FROM Mov M"
'                sql = sql & "   LEFT JOIN MovRef MR ON MR.nMovNro = M.nMovNro"
'                sql = sql & "   LEFT JOIN MovComprobanteReg MComp ON MComp.nMovNro = MR.nMovNroRef"
'                sql = sql & "   INNER JOIN MovGasto MG ON MG.nMovNro = M.nMovNro"
'                sql = sql & "   INNER JOIN Persona P ON P.cPersCod = MG.cPersCod"
'                sql = sql & "   LEFT JOIN Proveedor Prov ON Prov.cPersCod = P.cPersCod"
'                sql = sql & "   LEFT JOIN Persona PMN ON PMN.cPersCod = Prov.cPersCodBancoMN"
'                sql = sql & "   LEFT JOIN Persona PME ON PME.cPersCod = Prov.cPersCodBancoME" 'Descomentar Luego
                sql = sql & " )C ON C.nMovNro = M.nMovNro" 'END PASI
                          
            End If
            'END EJVG *******
        ElseIf psTipoInterfaz = "lbReporte" Then
            sql = sql & "  LEFT JOIN MovControlPagoSunat ps ON ps.nmovnro = m.nMovNro and ps.bvigente=1 and bvalido=1 "
        End If
        
        If bConsiderarDetra = 2 Then
            sql = sql & " JOIN "
            sql = sql & " ( Select C.nMovOrigen "
            sql = sql & "   From"
            sql = sql & "     ( Select C1.nMovNro as nMovOrigen, C2.nMovNro AS nMovDestino FROM"
            sql = sql & " ( select distinct mK.nMovNro "
            sql = sql & " from MovCta MK Inner Join Mov MV on MK.nmovnro=mv.nmovnro"
            sql = sql & " where MK.cCtaContCod like '" & cCtaDetraccion & "' and MK.nMovImporte <0 and MV.nmovflag=0 "
            sql = sql & " ) C1 "
            sql = sql & "       Left Join"
            sql = sql & " (Select MR.nMovNro, MR.nMovNroRef "
            sql = sql & " From MovRef MR Inner Join Mov MV1 on MR.nmovnroref=MV1.nmovnro "
            sql = sql & " Inner Join Mov MV2 on MR.nmovnro = mv2.nMovNro "
            sql = sql & " Where MV1.nMovFlag = 0 And mv2.nMovFlag = 0 and (isnull(mr.cAgeCodRef,'') = '' )  ) C2 " 'and mv1.copecod not in ('" & OpeRegPenalidadMN & "','" & OpeRegPenalidadME & "') 'and mv2.copecod not in ('" & OpeRegPenalidadMN & "','" & OpeRegPenalidadME & "')
            sql = sql & "       On C1.nMovNro=C2.nMovNroRef ) C"
            sql = sql & "   Group By C.nMovOrigen having Count(C.nMovDestino)=0"
            'ALPA 20090318*******************************************
            'sql = sql & " ) CC on M.nMovNro=CC.nMovOrigen"
            sql = sql & " ) CC on M.nMovNro=CC.nMovOrigen and isnull(C1.nTipoPago,0)=isnull(C2.nTipoPago,0) "
            '********************************************************
            
        End If
        sql = sql & " WHERE   m.nMovEstado = " & gMovEstContabMovContable & " and m.nMovFlag =0 "
        
        If psTipoInterfaz = "lbReporte" Then
             sql = sql & " and  ps.cmovnro is null and md.ndoctpo<>40 "
        End If
        
        '**************************************   D E T R A C C I O N   **************************************
        
        If bConsiderarDetra = 2 Then 'Busqueda para ejecutar la detraccion
             sql = sql & " and M.nMovNro not in( select M.nMovNro From Mov M Inner Join MovDetra MD On M.nMovNro=MD.nMovNro Where M.nMovFlag=0 and (MD.nEstado=2 or md.nestado=3)) "
        
        ElseIf bConsiderarDetra = 3 Then 'BUsqueda para Pago de Proveedores
            sql = sql & " and M.nMovNro not in( select M.nMovNro From Mov M Inner Join MovDetra MD On M.nMovNro=MD.nMovNro Where M.nMovFlag=0 and md.nestado=3) "
            sql = sql & " and (( "
            'sql = sql & "   M.nMovNro not in ( Select nMovOrigen from " & sTabla1 & ")"
            'sql = sql & "  or (M.nMovNro not in( Select nMovNro from " & sTabla2 & ")))) "
            
            sql = sql & "   M.nMovNro not in ( " & sqltable & ")"
            sql = sql & "  or (M.nMovNro not in( Select nMovNro from " & sTabla2 & ")))) "
        Else
            sql = sql & " and ref.nMovNro IS NULL "
        End If
        
        '**************************************   F I N   D E T R A C C I O N   **************************************
        If psTipoInterfaz = "PAGO" Then
            'EJVG20131119 ***
            'sql = sql & " and ps.cmovnro='" & psArchivo & "'"
            If Not pbPagoNuevo Then
                sql = sql & " and ps.cmovnro='" & psArchivo & "'"
            Else
                sql = sql & " and (md.dDocFecha>='" & Format(pdFechaEmisionIni, "yyyymmdd") & "' and md.dDocFecha<'" & Format(ldFechaEmisionFin, "yyyymmdd") & "') AND ISNULL(CASE WHEN SUBSTRING(mc.cCtaContCod,3,1)='1' THEN C.cPersCodBancoMN ELSE C.cPersCodBancoME END,'') IN (SELECT VALOR FROM dbo.fnc_getTblValoresTexto('" & psListaIFiCod & "'+',')) "
            End If
            'END EJVG *******
        Else
            sql = sql & "  and substring(m.cMovNro,1,8) >= '" & Format(pdFechaIni, "yyyymmdd") & " ' AND substring(m.cMovNro,1,8) <= '" & Format(pdFechaFin + 1, "yyyymmdd") & " '  "
            'sql = sql & "  and md.dDocFecha >= '" & Format(pdFechaIni, gsFormatoFecha) & "' AND md.dDocFecha < '" & Format(pdFechaFin + 1, gsFormatoFecha) & "' "

            
        End If
        sql = sql & " and copecod like '7%' and mc.cCtaContCod IN (" & psCtasCod & ")  " 'and m.copecod not in ('" & OpeRegPenalidadMN & "','" & OpeRegPenalidadME & "')
        'sql = sql & " AND ISNULL(ME.nMovMeImporte,mc.nMovImporte)<0" 'EJVG20131209 Para que no muestre reversiones de Acta de Conformidad Logistica
        '** Modify GITU 29-09-2008 ************************************'
        '** se cambio el order by del select para que siempre ordene **'
        '** del monto menor a mayor antes por nombre y nmovnro       **'
        'ALPA 20090318**************************************************
        'sql = sql & " and ref.nmovnro is null and isnull(mc.nTipoPago,0)=0 ORDER BY Prov.cpersnombre, MontoPago" 'm.nMovNro
        sql = sql & " and ref.nmovnro is null and isnull(mc.nTipoPago,0)= " & pnTipoPago & " ORDER BY Prov.cpersnombre, MontoPago" 'm.nMovNro
        '**************************************************************'
' and ref.nmovnro is null
Set GetDatosProvisionesProveedores = oConect.CargaRecordSet(sql)
       ' sql = " Drop table " & sTabla1
       ' oConect.Ejecutar sql
        sql = " Drop table " & sTabla2
        oConect.Ejecutar sql
       ' sql = " Drop table " & sTabla3
       ' oConect.Ejecutar sql
        sql = " Drop table " & sTabla4
        oConect.Ejecutar sql

Exit Function
GetDatosProvisionesProveedoresErr:
    Call RaiseError(MyUnhandledError, "DCajaGeneral:GetDatosProvisionesProveedores")
End Function


'Public Function GetDatosProvisionesProveedoresNuevo(psCtasCod As String, psDocs As String, pdFechaIni As Date, pdFechaFin As Date, Optional pbPend As Boolean = False, Optional bConsiderarDetra As Integer = 0, Optional cCtaDetraccion As String = "29180799", Optional psCtasCodME As String = "", Optional psTipoInterfaz As String = "", Optional psArchivo As String = "", Optional psTipo As String = "", Optional psOpeCod As String) As ADODB.Recordset
'Dim sql As String
'Dim nTieneDetra As New DMov
'Dim cCodOpe As String
'
'On Error GoTo GetDatosProvisionesProveedoresErr
'
'
'    psCtasCod = psCtasCod & IIf(psCtasCodME = "", "", "," & psCtasCodME)
'
'    If Mid(psOpeCod, 3, 1) = 2 Then
'        cCodOpe = "'702401' ,'702402' ,'702403'"
'    Else
'        cCodOpe = "'701401' ,'701402' ,'701403'"
'    End If
'
'    sql = "SELECT distinct md.dDocFecha, doc.cDocAbrev,mc.nMovImporte * -1 as nMovImporte ,md.nDocTpo, md.cDocNro, Prov.cpersnombre cPersona, Prov.nPersPersoneria, M.cMovDesc, Prov.cPersCod, "
'    sql = sql & "       m.cMovNro, m.nMovNro, mc.cCtaContCod"
'''''    If psTipoInterfaz = "PAGO" Or psTipoInterfaz = "lbReporte" Then
'''''        sql = sql & "   ,isnull(ps.nimportecoactivo,0) nImporteCoactivo, case when me.nMovMeImporte is NULL then (mc.nMovImporte*-1-ps.nimportecoactivo) else Round((me.nMovMeImporte*-1)-(ps.nimportecoactivo/ps.ntpocambio),2) end MontoPago, ps.nTpoCambio"
'''''    End If
'
'        sql = sql & " FROM   Mov m JOIN MovDoc md ON md.nMovNro = m.nMovNro "
'        sql = sql & "             JOIN MovCta mc ON mc.nMovNro = m.nMovNro LEFT JOIN MovMe ME ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem "
'        sql = sql & "             JOIN MovGasto mg ON mg.nMovNro = m.nMovNro "
'        sql = sql & "       LEFT JOIN (SELECT mr.nMovNro, mr.nMovNroRef FROM MovRef mr JOIN Mov m1 ON m1.nMovNro = mr.nMovNro "
'        sql = sql & "                   WHERE m1.nMovEstado = " & gMovEstContabMovContable & " and m1.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') and RTRIM(ISNULL(mr.cAgeCodRef,'')) = '' "
'        sql = sql & "                  ) ref ON  ref.nMovNroRef = m.nMovNro "
'        sql = sql & "             JOIN Persona Prov  ON Prov.cPersCod = mg.cPersCod "
'        sql = sql & "             left join Persid ProvI  ON ProvI.cPersCod = mg.cPersCod and cPersIdTpo=2"
'        sql = sql & "             JOIN Documento Doc ON Doc.nDocTpo = md.nDocTpo "
'        sql = sql & "  "
'
'
'        If bConsiderarDetra = 2 Then
'            sql = sql & " JOIN "
'            sql = sql & " ( Select C.nMovOrigen "
'            sql = sql & "   From"
'            sql = sql & "     ( Select C1.nMovNro as nMovOrigen, C2.nMovNro AS nMovDestino FROM"
'            sql = sql & " ( select distinct mK.nMovNro "
'            sql = sql & " from MovCta MK Inner Join Mov MV on MK.nmovnro=mv.nmovnro"
'            sql = sql & " where MK.cCtaContCod like '" & cCtaDetraccion & "' and MK.nMovImporte <0 and MV.nmovflag=0 "
'            sql = sql & " ) C1 "
'            sql = sql & "       Left Join"
'            sql = sql & " (Select MR.nMovNro, MR.nMovNroRef "
'            sql = sql & " From MovRef MR Inner Join Mov MV1 on MR.nmovnroref=MV1.nmovnro "
'            sql = sql & " Inner Join Mov MV2 on MR.nmovnro = mv2.nMovNro "
'            sql = sql & " Where MV1.nMovFlag = 0 AND MV2.cOpeCod not like '42[12]1%' And mv2.nMovFlag = 0 and (isnull(mr.cAgeCodRef,'') = '' )) C2 "
'            sql = sql & "       On C1.nMovNro=C2.nMovNroRef ) C"
'            sql = sql & "   Group By C.nMovOrigen having Count(C.nMovDestino)=0"
'            sql = sql & " ) CC on M.nMovNro=CC.nMovOrigen"
'
'        End If
'
'
'        sql = sql & " WHERE  m.cOpecod in (" & cCodOpe & ") and m.nMovEstado = " & gMovEstContabMovContable & " and m.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') "
'        If psTipoInterfaz = "lbReporte" Then
'             sql = sql & " and  ps.cmovnro is null and md.ndoctpo<>40  "
'        End If
'
'        '**************************************   D E T R A C C I O N   **************************************
'
'        If bConsiderarDetra = 2 Then 'Busqueda para ejecutar la detraccion
'
'            sql = sql & " and M.nMovNro not in( select M.nMovNro From Mov M Inner Join MovDetra MD On M.nMovNro=MD.nMovNro Where M.nMovFlag=0 and (MD.nEstado=2 or md.nestado=3)) "
'
'        ElseIf bConsiderarDetra = 3 Then 'BUsqueda para Pago de Proveedores
'
'            sql = sql & " and M.nMovNro not in( select M.nMovNro From Mov M Inner Join MovDetra MD On M.nMovNro=MD.nMovNro Where M.nMovFlag=0 and md.nestado=3) "
'
'            sql = sql & " and (( "
'
'            sql = sql & "   M.nMovNro not in"
'            sql = sql & " ( Select C.nMovOrigen "
'            sql = sql & "   From"
'            sql = sql & "     ( Select C1.nMovNro as nMovOrigen, C2.nMovNro AS nMovDestino FROM"
'                sql = sql & " ( select distinct mK.nMovNro "
'                sql = sql & " from MovCta MK Inner Join Mov MV on MK.nmovnro=Mv.nmovnro "
'                sql = sql & " where MK.cCtaContCod Like '" & cCtaDetraccion & "' and MK.nMovImporte <0 and MV.nmovflag=0 "
'                sql = sql & " ) C1 "
'            sql = sql & "       Left Join"
'                sql = sql & " (Select MR.nMovNro, MR.nMovNroRef "
'                sql = sql & " From MovRef MR Inner Join Mov MV1 on MR.nmovnroref = MV1.nmovnro And IsNull(MR.cAgeCodRef,'') = '' ANd MV1.nMovFlag = 0 ANd MV1.nMovEstado = 10  "
'                sql = sql & " Inner Join Mov MV2 on MR.nmovnro = mv2.nMovNro And mv2.nMovFlag = 0 And mv2.nMovEstado = 10  "
'                sql = sql & "  ) C2 "
'            sql = sql & "       On C1.nMovNro=C2.nMovNroRef   where  C2.nMovNroRef is not null ) C"
'            sql = sql & "   Group By C.nMovOrigen having Count(C.nMovDestino)<>1"
'
'            sql = sql & "   Union "
'            sql = sql & "   Select Distinct MR1.nMovNroref"
'            sql = sql & "   From MovRef MR1 "
'
'                sql = sql & " Inner Join Mov MV1 on MR1.nmovnroref=MV1.nmovnro And IsNull(MR1.cAgeCodRef,'') = '' ANd MV1.nMovFlag = 0 ANd MV1.nMovEstado = 10 " '
'                sql = sql & " Inner Join Mov MV2 on MR1.nmovnro = mv2.nMovNro And mv2.nMovFlag = 0 And mv2.nMovEstado = 10  " '
'
'            sql = sql & "   Inner Join MovDoc MD on MR1.nMovNroRef=MD.nMovNro"
'            sql = sql & "   where Len(MR1.nMovNro) > 0 "
'
'                sql = sql & " and MV1.nMovFlag = 0 And mv2.nMovFlag = 0 "
'
'            sql = sql & " and MR1.nMovNroref not in ("
'            sql = sql & "   select distinct movcta.nMovNro"
'            sql = sql & "   From MovCta  Inner Join Mov On Movcta.nMovNro= Mov.nMovNro "
'            sql = sql & "   where cCtaContCod = '29180799' and nMovImporte < 0 and Mov.nMovFlag=0 "
'            sql = sql & "   )"
'
'            sql = sql & " )) or (M.nMovNro in(select M.nMovNro From Mov M Inner Join MovDetra MD On M.nMovNro=MD.nMovNro Where M.nMovFlag=0 and MD.nEstado=2))) "
'
'        Else
'            sql = sql & " and ref.nMovNro IS NULL "
'        End If
'        '**************************************   F I N   D E T R A C C I O N   **************************************
'        'If psTipoInterfaz = "RECHAZO" Or psTipoInterfaz = "lbReporte" Then
'
'        If psTipoInterfaz = "PAGO" Then
'            sql = sql & " and ps.cmovnro='" & psArchivo & "' "
'        Else
'            sql = sql & "  and md.dDocFecha BETWEEN '" & Format(pdFechaIni, gsFormatoFecha) & "' AND '" & Format(pdFechaFin, gsFormatoFecha) & "' "
'        End If
'        sql = sql & "    and md.nDocTpo IN (" & psDocs & ") and mc.cCtaContCod IN (" & psCtasCod & ")  "
'        sql = sql & " ORDER BY Prov.cpersnombre, m.nMovNro"
'        'and cOpeCod not in (" & OpeRegPenalidadMN & " , " & OpeRegPenalidadMe & ")
'
'
'
'Set GetDatosProvisionesProveedoresNuevo = oConect.CargaRecordSet(sql)
'Exit Function
'GetDatosProvisionesProveedoresErr:
'    Call RaiseError(MyUnhandledError, "DCajaGeneral:GetDatosProvisionesProveedores")
'End Function



Public Function GetDatosProvisionesProveedoresNuevo(psCtasCod As String, psDocs As String, pdFechaIni As Date, pdFechaFin As Date, Optional pbPend As Boolean = False, Optional bConsiderarDetra As Integer = 0, Optional cCtaDetraccion As String = "29180799", Optional psCtasCodME As String = "", Optional psTipoInterfaz As String = "", Optional psArchivo As String = "", Optional psTipo As String = "", Optional psOpecod As String) As ADODB.Recordset
Dim sql As String
Dim nTieneDetra As New DMov
Dim cCodOpe As String
Dim lsCodOpe As String

On Error GoTo GetDatosProvisionesProveedoresErr
    
    
    psCtasCod = psCtasCod & IIf(psCtasCodME = "", "", "," & psCtasCodME)
    
    If Mid(psOpecod, 3, 1) = 2 Then
        cCodOpe = "'702401' ,'702402' ,'702403','702404','702405','702406'"
    Else
        cCodOpe = "'701401' ,'701402' ,'701403','701404','701405','701406'"
    End If
    
    
    Select Case psOpecod 'Operacion que no incluye en la busqueda
           Case 401581
                lsCodOpe = "401582,401585" 'Codigo de Operacion de Embargo y Fiel Cumplimiento MN
           Case 401582
                lsCodOpe = "401581,401585" 'Codigo de Operacion de Detraccion y Fiel Cumplimiento MN
           Case 402581
                lsCodOpe = "402582,402585" 'Codigo de Operacion de Embargo y Fiel Cumplimiento ME
           Case 402582
                lsCodOpe = "402581,402585" 'Codigo de Operacion de Detraccion y Fiel Cumplimiento ME
           
           Case 401585
                lsCodOpe = "401581,401582" 'Codigo de Operacion de Detraccion y Fiel Cumplimiento MN
           Case 402585
                lsCodOpe = "402581,402582" 'Codigo de Operacion de Detraccion y Fiel Cumplimiento ME
           
           
    End Select
    
    
   
    sql = "SELECT distinct md.dDocFecha, doc.cDocAbrev,mc.nMovImporte * -1 as nMovImporte ,md.nDocTpo, md.cDocNro, Prov.cpersnombre cPersona, Prov.nPersPersoneria, M.cMovDesc, Prov.cPersCod, "
    sql = sql & "       m.cMovNro, m.nMovNro, mc.cCtaContCod,mE.nMovMEImporte * -1 as nMovImporteME"
''''    If psTipoInterfaz = "PAGO" Or psTipoInterfaz = "lbReporte" Then
''''        sql = sql & "   ,isnull(ps.nimportecoactivo,0) nImporteCoactivo, case when me.nMovMeImporte is NULL then (mc.nMovImporte*-1-ps.nimportecoactivo) else Round((me.nMovMeImporte*-1)-(ps.nimportecoactivo/ps.ntpocambio),2) end MontoPago, ps.nTpoCambio"
''''    End If
        
        sql = sql & " FROM   Mov m JOIN MovDoc md ON md.nMovNro = m.nMovNro "
        sql = sql & "             JOIN MovCta mc ON mc.nMovNro = m.nMovNro LEFT JOIN MovMe ME ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem "
        sql = sql & "             JOIN MovGasto mg ON mg.nMovNro = m.nMovNro "
        sql = sql & "       LEFT JOIN (SELECT mr.nMovNro, mr.nMovNroRef FROM MovRef mr JOIN Mov m1 ON m1.nMovNro = mr.nMovNro "
        sql = sql & "                   WHERE m1.nMovEstado = " & gMovEstContabMovContable & " and m1.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') and RTRIM(ISNULL(mr.cAgeCodRef,'')) = '' "
        sql = sql & "                  ) ref ON  ref.nMovNroRef = m.nMovNro "
        sql = sql & "             JOIN Persona Prov  ON Prov.cPersCod = mg.cPersCod "
        sql = sql & "             left join Persid ProvI  ON ProvI.cPersCod = mg.cPersCod and cPersIdTpo=2"
        sql = sql & "             JOIN Documento Doc ON Doc.nDocTpo = md.nDocTpo "
        sql = sql & "  "

               
        If bConsiderarDetra = 2 Then
            sql = sql & " JOIN "
            sql = sql & " ( Select C.nMovOrigen "
            sql = sql & "   From"
            sql = sql & "     ( Select C1.nMovNro as nMovOrigen, C2.nMovNro AS nMovDestino FROM"
            sql = sql & " ( select distinct mK.nMovNro "
            sql = sql & " from MovCta MK Inner Join Mov MV on MK.nmovnro=mv.nmovnro"
            sql = sql & " where MK.cCtaContCod like '" & cCtaDetraccion & "' and MK.nMovImporte <0 and MV.nmovflag=0 "
            sql = sql & " ) C1 "
            sql = sql & "       Left Join"
            sql = sql & " (Select MR.nMovNro, MR.nMovNroRef "
            sql = sql & " From MovRef MR Inner Join Mov MV1 on MR.nmovnroref=MV1.nmovnro "
            sql = sql & " Inner Join Mov MV2 on MR.nmovnro = mv2.nMovNro "
            sql = sql & " Where MV1.nMovFlag = 0 AND MV2.cOpeCod not like '42[12]1%' and MV2.cOpeCod not in (" & lsCodOpe & ") And mv2.nMovFlag = 0 and (isnull(mr.cAgeCodRef,'') = '' )) C2 "
            sql = sql & "       On C1.nMovNro=C2.nMovNroRef ) C"
            sql = sql & "   Group By C.nMovOrigen having Count(C.nMovDestino)=0"
            sql = sql & " ) CC on M.nMovNro=CC.nMovOrigen"
            
        End If
        
        
        sql = sql & " WHERE  m.cOpecod in (" & cCodOpe & ") and m.nMovEstado = " & gMovEstContabMovContable & " and m.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') "
        If psTipoInterfaz = "lbReporte" Then
             sql = sql & " and  ps.cmovnro is null and md.ndoctpo<>40  "
        End If
        
        '**************************************   D E T R A C C I O N   **************************************
        
        If bConsiderarDetra = 2 Then 'Busqueda para ejecutar la detraccion
        
            sql = sql & " and M.nMovNro not in( select M.nMovNro From Mov M Inner Join MovDetra MD On M.nMovNro=MD.nMovNro Where M.nMovFlag=0 and (MD.nEstado=2 or md.nestado=3)) "
        
        ElseIf bConsiderarDetra = 3 Then 'BUsqueda para Pago de Proveedores
            
            sql = sql & " and M.nMovNro not in( select M.nMovNro From Mov M Inner Join MovDetra MD On M.nMovNro=MD.nMovNro Where M.nMovFlag=0 and md.nestado=3) "
            
            sql = sql & " and (( "
             
            sql = sql & "   M.nMovNro not in"
            sql = sql & " ( Select C.nMovOrigen "
            sql = sql & "   From"
            sql = sql & "     ( Select C1.nMovNro as nMovOrigen, C2.nMovNro AS nMovDestino FROM"
                sql = sql & " ( select distinct mK.nMovNro "
                sql = sql & " from MovCta MK Inner Join Mov MV on MK.nmovnro=Mv.nmovnro "
                sql = sql & " where MK.cCtaContCod Like '" & cCtaDetraccion & "' and MK.nMovImporte <0 and MV.nmovflag=0 "
                sql = sql & " ) C1 "
            sql = sql & "       Left Join"
                sql = sql & " (Select MR.nMovNro, MR.nMovNroRef "
                sql = sql & " From MovRef MR Inner Join Mov MV1 on MR.nmovnroref = MV1.nmovnro And IsNull(MR.cAgeCodRef,'') = '' ANd MV1.nMovFlag = 0 ANd MV1.nMovEstado = 10  "
                sql = sql & " Inner Join Mov MV2 on MR.nmovnro = mv2.nMovNro And mv2.nMovFlag = 0 And mv2.nMovEstado = 10  "
                sql = sql & "  ) C2 "
            sql = sql & "       On C1.nMovNro=C2.nMovNroRef   where  C2.nMovNroRef is not null ) C"
            sql = sql & "   Group By C.nMovOrigen having Count(C.nMovDestino)<>1"
            
            sql = sql & "   Union "
            sql = sql & "   Select Distinct MR1.nMovNroref"
            sql = sql & "   From MovRef MR1 "
                
                sql = sql & " Inner Join Mov MV1 on MR1.nmovnroref=MV1.nmovnro And IsNull(MR1.cAgeCodRef,'') = '' ANd MV1.nMovFlag = 0 ANd MV1.nMovEstado = 10 " '
                sql = sql & " Inner Join Mov MV2 on MR1.nmovnro = mv2.nMovNro And mv2.nMovFlag = 0 And mv2.nMovEstado = 10  " '
            
            sql = sql & "   Inner Join MovDoc MD on MR1.nMovNroRef=MD.nMovNro"
            sql = sql & "   where Len(MR1.nMovNro) > 0 "
            
                sql = sql & " and MV1.nMovFlag = 0 And mv2.nMovFlag = 0 "
            
            sql = sql & " and MR1.nMovNroref not in ("
            sql = sql & "   select distinct movcta.nMovNro"
            sql = sql & "   From MovCta  Inner Join Mov On Movcta.nMovNro= Mov.nMovNro "
            sql = sql & "   where cCtaContCod = '29180799' and nMovImporte < 0 and Mov.nMovFlag=0 "
            sql = sql & "   )"
            
            sql = sql & " )) or (M.nMovNro in(select M.nMovNro From Mov M Inner Join MovDetra MD On M.nMovNro=MD.nMovNro Where M.nMovFlag=0 and MD.nEstado=2))) "
         
        Else
            sql = sql & " and ref.nMovNro IS NULL "
        End If
        '**************************************   F I N   D E T R A C C I O N   **************************************
        'If psTipoInterfaz = "RECHAZO" Or psTipoInterfaz = "lbReporte" Then
            
        If psTipoInterfaz = "PAGO" Then
            sql = sql & " and ps.cmovnro='" & psArchivo & "' "
        Else
            sql = sql & "  and md.dDocFecha BETWEEN '" & Format(pdFechaIni, gsFormatoFecha) & "' AND '" & Format(pdFechaFin, gsFormatoFecha) & "' "
        End If
        sql = sql & "    and md.nDocTpo IN (" & psDocs & ") and mc.cCtaContCod IN (" & psCtasCod & ")  "
        sql = sql & " ORDER BY Prov.cpersnombre, m.nMovNro"
        'and cOpeCod not in (" & OpeRegPenalidadMN & " , " & OpeRegPenalidadMe & ")
    
    

Set GetDatosProvisionesProveedoresNuevo = oConect.CargaRecordSet(sql)
Exit Function
GetDatosProvisionesProveedoresErr:
    Call RaiseError(MyUnhandledError, "DCajaGeneral:GetDatosProvisionesProveedores")
End Function

Public Function GrabaPagoProvSUNAT(psMovNro As String, psOpecod As String, psMovDesc As String, _
                                   psCtaContDebeBS As String, psCtaContDebeSS As String, psCtaContDebeRHS As String, psCtaContDebeRHJS As String, psCtaContDebeSEGUS As String, psCtaContDebePAGOVARIOSS As String, _
                                   psCtaContDebeBD As String, psCtaContDebeSD As String, psCtaContDebeRHD As String, psCtaContDebeRHJD As String, psCtaContDebeSEGUD As String, psCtaContDebePAGOVARIOSD As String, _
                                   psCtaContHaber As String, _
                                   pnImporteBS As Currency, pnImporteSS As Currency, pnImporteRHS As Currency, pnImporteRHJS As Currency, pnImporteSEGUS As Currency, pnImportePAGOVARIOSS As Currency, _
                                   pnImporteBD As Currency, pnImporteSD As Currency, pnImporteRHD As Currency, pnImporteRHJD As Currency, pnImporteSEGUD As Currency, pnImportePAGOVARIOSD As Currency, _
                                   psPersCod As String, psTpoIF As String, psPersCodIf As String, psCtaIf As String, _
                                   prsBilletaje As ADODB.Recordset, psDocTpo As String, psDocNro As String, psFechaDoc As String, psDocVoucher As String, rsDocs As ADODB.Recordset, psCuentaAho As String, _
                                   ByVal psCtaDiferencia As String, ByVal pnMontoDiferencia As Currency, Optional ByVal gbBitCentral As Boolean, Optional pbGrabaOpeNegocio As Boolean = False, Optional pnRetAct As Currency = 0, Optional pnITFValor As Double, _
                                   Optional ByVal bDetraccion As Boolean = False, Optional pnAjustarMonto As Currency = 0, Optional pnAjustarAB As Currency = 0, Optional pnAjustarAS As Currency = 0, Optional pnAjustarARH As Currency = 0, Optional pnAjustarARHJ As Currency = 0, _
                                   Optional pnAjustarASEGU As Currency = 0, Optional pnAjustarAPAGOVARIOS As Currency = 0, _
                                   Optional lsCtaITFD As String, Optional lsCtaITFH As String, Optional nTasaITF As Double = 0, Optional bAfectoITF As Boolean = True) As Integer
Dim lnMovItem  As Long
Dim lnMovOrden As Integer
Dim lnMovNro As Long
Dim oDMov    As DMov
Dim lbTrans  As Boolean
Dim lsMovNro As String
Dim lsCadBol As String
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
Dim sSql As String
Dim lsMsgErr As String

On Error GoTo GrabaPagoProveedorErr
GetTipCambio GetFechaMov(psMovNro, True)
If Mid(psCtaContHaber, 3, 1) = Moneda.gMonedaExtranjera Then
    lMN = False
Else
    lMN = True
End If
lbBitReten = IIf(oConst.LeeConstSistema(gConstSistBitRetencion6Porcent) = 1, True, False)
lsRetDocNro = oFun.GeneraDocNro(TpoDocAlmacenComprdeRetencion6)
If lbBitReten Then
    lsCtaReten = oConst.LeeConstSistema(gConstSistCtaRetencion6Porcent)
End If
oDMov.BeginTrans
lbTrans = True
GrabaPagoProvSUNAT = 1

oDMov.InsertaMov psMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporteBS + pnImporteSS + pnImporteRHS + pnImporteRHJS + pnImporteSEGUS + pnImportePAGOVARIOSS + pnImporteBD + pnImporteSD + pnImporteRHD + pnImporteRHJD + pnImporteSEGUD + pnImportePAGOVARIOSD, "0", "0"
If bDetraccion = True Then
    oDMov.InsertaMovDetra lnMovNro, , Right(psMovNro, 4)
End If
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
If pnImporteBS <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeBS, pnImporteBS
End If
If pnImporteBD <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeBD, pnImporteBD
   
   If psOpecod = "421120" Then
        oDMov.InsertaMovMe lnMovNro, lnMovItem, pnAjustarAB
   End If
End If
If pnImporteSS <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeSS, pnImporteSS
End If
If pnImporteSD <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeSD, pnImporteSD
   If psOpecod = "421120" Then
        oDMov.InsertaMovMe lnMovNro, lnMovItem, pnAjustarAS
   End If
End If

If pnImporteRHS <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeRHS, pnImporteRHS
End If
If pnImporteRHD <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeRHD, pnImporteRHD
   If psOpecod = "421120" Then
        oDMov.InsertaMovMe lnMovNro, lnMovItem, pnAjustarARH
   End If
End If

If pnImporteRHJS <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeRHJS, pnImporteRHJS
End If
If pnImporteRHJD <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeRHJD, pnImporteRHJD
   If psOpecod = "421120" Then
        oDMov.InsertaMovMe lnMovNro, lnMovItem, pnAjustarARHJ
   End If
End If

If pnImporteSEGUS <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeSEGUS, pnImporteSEGUS
End If
If pnImporteSEGUD <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebeSEGUD, pnImporteSEGUD
   If psOpecod = "421120" Then
        oDMov.InsertaMovMe lnMovNro, lnMovItem, pnAjustarASEGU
   End If
End If

If pnImportePAGOVARIOSS <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebePAGOVARIOSS, pnImportePAGOVARIOSS
End If
If pnImportePAGOVARIOSD <> 0 Then
   lnMovItem = lnMovItem + 1
   oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebePAGOVARIOSD, pnImportePAGOVARIOSD
   If psOpecod = "421120" Then
        oDMov.InsertaMovMe lnMovNro, lnMovItem, pnAjustarAPAGOVARIOS
   End If
End If




lnRetAct = 0
lnRetActME = 0
If lbBitReten Then
   lnRetAct = pnRetAct
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
    If (pnImporteBS + pnImporteSS + pnImporteRHS + pnImporteRHJS + pnImporteSEGUS + pnImportePAGOVARIOSS + pnImporteBD + pnImporteSD + pnImporteRHD + pnImporteRHJD + pnImporteSEGUD + pnImportePAGOVARIOSD - IIf(lMN, lnRetAct, lnRetActME)) <> 0 Then
        lnMovItem = lnMovItem + 1: lnMovOrden = 0
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber, (pnImporteBS + pnImporteSS + pnImporteRHS + pnImporteRHJS + pnImporteSEGUS + pnImportePAGOVARIOSS + pnImporteBD + pnImporteSD + pnImporteRHD + pnImporteRHJD + pnImporteSEGUD + pnImportePAGOVARIOSD - IIf(lnRetAct + lnRetActME > 0, IIf(lMN, lnRetAct, lnRetActME), 0)) * -1
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
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDiferencia, pnMontoDiferencia * -1
End If
'Referenciamos todos los Documentos que se Pagan
If Not rsDocs Is Nothing Then
   Do While Not rsDocs.EOF
      If rsDocs!Itm = "1" Then
         oDMov.InsertaMovRef lnMovNro, rsDocs!nMovNro
         oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeBS, 0, False
         oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeSS, 0, False
         oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeBD, 0, False
         oDMov.ActualizaMovPendientesRend rsDocs!nMovNro, psCtaContDebeSD, 0, False
         oDMov.ActualizaMovSUNAT rsDocs!nMovNro, lnMovNro
'         If bDetraccion = False Then
'            oDMov.ActualizaMovDetra rsDocs!nMovNro, 3, , Right(psMovNro, 4)
'         End If
         
      End If
      rsDocs.MoveNext
   Loop
End If

'--Calculo del ITF
'--  JOHN
If bAfectoITF = True Then 'bit si esta afecto al ITF el pago a Proveedor
    'Dim lnMontoITF As Currency
    Dim lnMontoITF As Double '*** PEAC 20110331
    If lsCtaITFD <> "" And lsCtaITFH <> "" And nTasaITF > 0 Then
        lnMontoITF = fgTruncar((pnImporteBD + pnImporteSD + pnImporteRHD + pnImporteRHJD + pnImporteSEGUD + pnImportePAGOVARIOSD - lnRetAct) * nTasaITF, 2)
        If lnMontoITF > 0 Then
            lnMovItem = lnMovItem + 1
            oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFD, lnMontoITF
            lnMovItem = lnMovItem + 1
            oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFH, lnMontoITF * -1
        End If
    End If
End If
'----------


If Mid(psOpecod, 3, 1) = 2 Then
    If lbBitReten Then
        oDMov.GeneraMovME lnMovNro, psMovNro, gnTipCambioPonderado
    Else
        oDMov.GeneraMovME lnMovNro, psMovNro
    End If
    'pnAjustarMonto
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"
If psDocTpo = TpoDocNotaAbono Or psDocTpo = TpoDocNotaCargo Then
    'oDMov.InsertaRegProveedorAbonoCta lnMovNro, psDocNro, psCuentaAho
    
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
           lnSaldo = oCapta.CapAbonoCuentaAho(psCuentaAho, pnImporteBS + pnImporteSS - IIf(lMN, lnRetAct, lnRetActME), gCGPagProvAbonoCent, lsMovNroNegocio, psMovDesc, TpoDocNotaAbono, psDocNro, , , , , , , , , , True, False, , , oDMov.GetConexion, False)
        Else
            If oCon.AbreConexionRemota(Left(psCuentaAho, 2)) Then
                oCon.BeginTrans
                lbGraboNegocio = True
                If psDocTpo = TpoDocNotaAbono Then
                    lnSaldo = oDis.Abono(psCuentaAho, pnImporteBS + pnImporteSS - IIf(lMN, lnRetAct, lnRetActME), gCGPagProvAbonoDist, gCGPagProvAbonoDist, "112" & Left(psCuentaAho, 2), Right(psMovNro, 4), psDocNro, "PAGO : " & psMovDesc, oCon, GetFechaMov(psMovNro, True))
                End If
                oCon.CommitTrans
            
'                Dim oConFun As NContFunciones
'                Set oConFun = New NContFunciones
'                oConFun.SetNroNANC True, Mid(psMovNro, 18, 2)
            End If
            oCon.CierraConexion
            Set oCon = Nothing
            Set oDis = Nothing
        End If
    End If
    
'ElseIf psDocTpo = TpoDocCheque Then
'    InsertaNumeroCheque psPersCodIF, psCtaIf, psDocNro
End If
oDMov.CommitTrans
lbTrans = False
GrabaPagoProvSUNAT = 0
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

Public Function GetDatosPagoProveedoresSUNAT(pdFechaIni As Date, pdFechaFin As Date, psCodOpe As String, Optional psArchivo As String = "") As ADODB.Recordset
Dim sql         As String
Dim rs          As New ADODB.Recordset
On Error GoTo GetDatosPagoProveedoresSUNAT

    'Modificado PASIERS1242014
'    sql = " select distinct"
'    sql = sql & " md.dDocFecha,md.cDocNro, doc.cDocAbrev, md.nDocTpo,Prov.cpersnombre cPersona,"
'    sql = sql & " M.cMovDesc, Prov.cPersCod,m.cMovNro, m.nMovNro, mc.cCtaContCod,"
'    sql = sql & " ISNULL(me.nMovMeImporte,mc.nMovImporte) * -1 as nMovImporte, mc.nMovImporte * -1 as nMovImporteSoles"
'    sql = sql & " ,ps.nimportecoactivo, ref.cMovNro MovPago, ps.nTpoCambio ,0 Penalidad "
'    sql = sql & " from movControlPagoSunat ps"
'    sql = sql & " join mov m on ps.nMovNro=m.nMovNro"
'    sql = sql & " join movdoc md on m.nMovNro=md.nMovNro"
'    'sql = sql & " left join (select mr.nmovnroref from movref mr "
'    'sql = sql & "           join mov mm on mm.nMovNro=mr.nMovNro and  mm.copecod = '421116' and mm.nMovFlag = " & MovFlag.gMovFlagVigente & " and mm.nMovEstado = " & MovEstado.gMovEstContabMovContable & " )c on c.nmovnroref=ps.nmovnro"
'    sql = sql & " LEFT JOIN ( SELECT mr.nMovNro, mr.nMovNroRef, m1.cMovNro FROM MovRef mr JOIN Mov m1 ON m1.nMovNro = mr.nMovNro"
'    sql = sql & "             WHERE m1.nMovEstado = " & MovEstado.gMovEstContabMovContable & " and m1.nMovFlag = " & MovFlag.gMovFlagVigente & " and RTRIM(ISNULL(mr.cAgeCodRef,'')) = '' And m1.cOpeCod Like '42_11[12345]' ) ref ON  ref.nMovNroRef = m.nMovNro"
'    sql = sql & " inner join Documento Doc ON Doc.nDocTpo = md.nDocTpo"
'    sql = sql & " join MovGasto mg on m.nMovNro=mg.nMovNro"
'    sql = sql & " JOIN Persona Prov  ON Prov.cPersCod = mg.cPersCod"
'    'ALPA 20110912**************************
'    'sql = sql & " join MovCta mc on ps.nMovNro =mc.nMovNro and mc.cCtaContCod IN ('251601','25160202','252601','25260202')"
'    sql = sql & " join MovCta mc on ps.nMovNro =mc.nMovNro and mc.cCtaContCod IN ('251601','25160202','252601','25260202','25141915','25241915')"
'    '***************************************
'    sql = sql & " left join MovME me on m.nMovNro =me.nMovNro and mc.nmovitem = me.nmovitem "
'    sql = sql & " where cMovNroRes Is Not Null and nimportecoactivo>0 and ps.nMovPagadoSunat is null and m.nMovFlag = " & MovFlag.gMovFlagVigente & " and m.nMovEstado = " & MovEstado.gMovEstContabMovContable & " "
'    sql = sql & " and bvigente=1 and bvalido=1 "
'    sql = sql & " and md.dDocFecha BETWEEN '" & Format(pdFechaIni, gsFormatoFecha) & "' AND '" & Format(pdFechaFin, gsFormatoFecha) & "' and mc.nMovimporte < 0 AND ISNULL(MC.nTipoPago,0)=0"
    
    sql = "stp_sel_ERS1242014_GetDatosPagoProveedoresSUNAT '" & Format(pdFechaIni, "yyyyMMdd") & "','" & Format(pdFechaFin, "yyyyMMdd") & "'"
    
    Set GetDatosPagoProveedoresSUNAT = oConect.CargaRecordSet(sql)
    
    'END PASI
     'and ref.cMovNro is Not Null
    Exit Function
GetDatosPagoProveedoresSUNAT:
    MsgBox Err.Description, vbExclamation, "Aviso"
End Function

'EJVG20121018 ***
Public Sub dBeginTrans()
    oConect.BeginTrans
End Sub
Public Sub dRollbackTrans()
    oConect.RollbackTrans
    oConect.Ejecutar "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
End Sub
Public Sub dCommitTrans()
    oConect.CommitTrans
End Sub
Public Function RecuperaFlujoCajaProyectado() As ADODB.Recordset
    Dim sql As String
    On Error GoTo ErrRecuperaFlujoCajaProyectado
    sql = "Exec stp_sel_FlujoCajaProyectado"
    Set RecuperaFlujoCajaProyectado = oConect.CargaRecordSet(sql)
    Exit Function
ErrRecuperaFlujoCajaProyectado:
        Call RaiseError(MyUnhandledError, "DCajaGeneral:RecuperaFlujoCajaProyectado")
End Function
Public Function RecuperaFlujoCajaProyectadoxAnio(ByVal pnAnio As Integer, ByVal pnMoneda As Moneda) As ADODB.Recordset
    Dim sql As String
    On Error GoTo ErrRecuperaFlujoCajaProyectadoxAnio
    sql = "Exec stp_sel_FlujoCajaProyectadoxAnio " & pnAnio & "," & pnMoneda
    Set RecuperaFlujoCajaProyectadoxAnio = oConect.CargaRecordSet(sql)
    Exit Function
ErrRecuperaFlujoCajaProyectadoxAnio:
        Call RaiseError(MyUnhandledError, "DCajaGeneral:RecuperaFlujoCajaProyectadoxAnio")
End Function
Public Function obtenerNuevoCodFCP(ByVal pnAnio As Integer, ByVal pnMoneda As Moneda) As Integer
    Dim rs As New ADODB.Recordset
    Dim sql As String
    
    On Error GoTo ErrNuevoFCPCod
    sql = "select ISNULL(MAX(nFCPActual),0)+1 nCodNuevo from FlujoCajaProy C where C.nAnio=" & pnAnio & " And C.nMoneda=" & pnMoneda
    Set rs = oConect.CargaRecordSet(sql)
    obtenerNuevoCodFCP = rs!nCodNuevo
    
    Set rs = Nothing
    Exit Function
ErrNuevoFCPCod:
    Err.Raise Err.Number, "Error", Err.Description
End Function
Public Sub InsertaFlujoCajaProyectado(ByVal pnAnio As Integer, ByVal pnMoneda As Moneda, ByVal pnFCPActual As Integer, ByVal psUsuario As String, ByVal pdFecha As Date)
    Dim sql As String
    On Error GoTo ErrInsertaFCP
    sql = "Exec stp_ins_FlujoCajaProyectado " & pnAnio & "," & pnMoneda & "," & pnFCPActual & ",'" & psUsuario & "','" & Format(pdFecha, "yyyymmdd hh:mm:ss") & "'"
    oConect.Ejecutar sql
    Exit Sub
ErrInsertaFCP:
    Err.Raise Err.Number, "Error", Err.Description
End Sub
Public Sub ActualizaFlujoCajaProyectado(ByVal pnAnio As Integer, ByVal pnMoneda As Moneda, Optional ByVal pnFCPActual As Integer = 0, Optional ByVal psUsuario As String = "", Optional ByVal pdFecha As Date = "01/01/1900")
    Dim sql As String
    Dim lsCampos As String
    On Error GoTo ErrInsertaFCP
    
    If pnFCPActual <> 0 Then
        lsCampos = lsCampos & IIf(Len(lsCampos) > 0, " And ", "") & "nFCPActual = " & pnFCPActual
    End If
    If psUsuario <> "" Then
        lsCampos = lsCampos & IIf(Len(lsCampos) > 0, " And ", "") & "cUserProceso = " & psUsuario
    End If
    If DateDiff("d", "01/01/1900", pdFecha) > 0 Then
        lsCampos = lsCampos & IIf(Len(lsCampos) > 0, " And ", "") & " dFechaProceso = '" & Format(pdFecha, "yyyymmdd hh:mm:ss") & "'"
    End If
    
    sql = "Update FlujoCajaProy Set " & lsCampos & " where nAnio=" & pnAnio & " And nMoneda=" & pnMoneda
    oConect.Ejecutar sql
    Exit Sub
ErrInsertaFCP:
    Err.Raise Err.Number, "Error", Err.Description
End Sub
Public Sub InsertaFlujoCajaProyectadoDet(ByVal pnAnio As Integer, ByVal pnMoneda As Moneda, ByVal pnFCPActual As Integer, ByVal pnMes As Integer, ByVal pnFCPId As Integer, ByVal pnMonto As Currency)
    Dim sql As String
    On Error GoTo ErrInsertaFCPDet
    sql = "Exec stp_ins_FlujoCajaProyectadoDet " & pnAnio & "," & pnMoneda & "," & pnFCPActual & "," & pnMes & "," & pnFCPId & "," & pnMonto
    oConect.Ejecutar sql
    Exit Sub
ErrInsertaFCPDet:
    Err.Raise Err.Number, "Error", Err.Description
End Sub
'END EJVG *******
'FRHU20140118 RQ13825
Public Function RecuperaProyeccionAhorroAge() As ADODB.Recordset
    Dim sql As String
    On Error GoTo ErrRecuperaProyeccionAhorroAge
    sql = "Exec stp_sel_ProyeccionAhorroAge"
    Set RecuperaProyeccionAhorroAge = oConect.CargaRecordSet(sql)
    Exit Function
ErrRecuperaProyeccionAhorroAge:
        Call RaiseError(MyUnhandledError, "DCajaGeneral:RecuperaProyeccionAhorroAge")
End Function
Public Function RecuperaProyeccionAhorroAgexAnio(ByVal pnAnio As String) As ADODB.Recordset
    Dim sql As String
    On Error GoTo ErrRecuperaProyeccionAhorroAgexAnio
    sql = "Exec stp_sel_ProyeccionAhorroAgePorAnio '" & pnAnio & "'"
    Set RecuperaProyeccionAhorroAgexAnio = oConect.CargaRecordSet(sql)
    Exit Function
ErrRecuperaProyeccionAhorroAgexAnio:
        Call RaiseError(MyUnhandledError, "DCajaGeneral:RecuperaProyeccionAhorroAgexAnio")
End Function
Public Sub InsertaProyeccionAhorroAge(ByVal pnAnio As String, pnMes As String, ByVal pnAgeCod As String, ByVal pnMonto As Currency)
    Dim sql As String
    On Error GoTo ErrInsertaProyeccionAhorroAge
    sql = "Exec stp_ins_upd_ProyeccionAhorroAge '" & pnAnio & "','" & pnMes & "','" & pnAgeCod & "', " & pnMonto
    oConect.Ejecutar sql
    Exit Sub
ErrInsertaProyeccionAhorroAge:
    Err.Raise Err.Number, "Error", Err.Description
End Sub
'FIN FRHU20140118
'FRHU20140121 RQ13826
Public Function RecuperaRptAhorros(ByVal psAnioDesde As String, ByVal psMesDesde As String, ByVal psAnioHasta As String, ByVal psMesHasta As String, ByVal psAgeCod As String) As ADODB.Recordset
    Dim sql As String
    On Error GoTo ErrRecuperaRptAhorros
    sql = "Exec stp_sel_ReporteEvolucionDepositoAhorroCorriente '" & psAnioDesde & "','" & psMesDesde & "','" & psAnioHasta & "','" & psMesHasta & "','" & psAgeCod & "'"
    Set RecuperaRptAhorros = oConect.CargaRecordSet(sql)
    Exit Function
ErrRecuperaRptAhorros:
        Call RaiseError(MyUnhandledError, "DCajaGeneral:RecuperaRptAhorros")
End Function
Public Function RecuperaRptDepositoPlazoFijo(ByVal psAnioDesde As String, ByVal psMesDesde As String, ByVal psAnioHasta As String, ByVal psMesHasta As String, ByVal psAgeCod As String) As ADODB.Recordset
    Dim sql As String
    On Error GoTo ErrRecuperaRptDepositoPlazoFijo
    sql = "Exec stp_sel_ReporteEvolucionDepositoPlazoFijo '" & psAnioDesde & "','" & psMesDesde & "','" & psAnioHasta & "','" & psMesHasta & "','" & psAgeCod & "'"
    Set RecuperaRptDepositoPlazoFijo = oConect.CargaRecordSet(sql)
    Exit Function
ErrRecuperaRptDepositoPlazoFijo:
        Call RaiseError(MyUnhandledError, "DCajaGeneral:RecuperaRptDepositoPlazoFijo")
End Function
Public Function RecuperaRptDepositoCTS(ByVal psAnioDesde As String, ByVal psMesDesde As String, ByVal psAnioHasta As String, ByVal psMesHasta As String, ByVal psAgeCod As String) As ADODB.Recordset
    Dim sql As String
    On Error GoTo ErrRecuperaRptDepositoCTS
    sql = "Exec stp_sel_ReporteEvolucionDepositoCTS '" & psAnioDesde & "','" & psMesDesde & "','" & psAnioHasta & "','" & psMesHasta & "','" & psAgeCod & "'"
    Set RecuperaRptDepositoCTS = oConect.CargaRecordSet(sql)
    Exit Function
ErrRecuperaRptDepositoCTS:
        Call RaiseError(MyUnhandledError, "DCajaGeneral:RecuperaRptDepositoCTS")
End Function
Public Function RecuperaRptDepositoTOTAL(ByVal psAnioDesde As String, ByVal psMesDesde As String, ByVal psAnioHasta As String, ByVal psMesHasta As String, ByVal psAgeCod As String) As ADODB.Recordset
    Dim sql As String
    On Error GoTo ErrRecuperaRptDepositoTOTAL
    sql = "Exec stp_sel_ReporteEvolucionDepositoTotal '" & psAnioDesde & "','" & psMesDesde & "','" & psAnioHasta & "','" & psMesHasta & "','" & psAgeCod & "'"
    Set RecuperaRptDepositoTOTAL = oConect.CargaRecordSet(sql)
    Exit Function
ErrRecuperaRptDepositoTOTAL:
        Call RaiseError(MyUnhandledError, "DCajaGeneral:RecuperaRptDepositoTOTAL")
End Function
Public Function ReporteDepositosTotales(ByVal psAnioDesde As String, ByVal psMesDesde As String, ByVal psAnioHasta As String, ByVal psMesHasta As String, ByVal psAgeCod As String) As ADODB.Recordset
    Dim sql As String
    On Error GoTo ErrReporteDepositosTotales
    sql = "Exec stp_sel_ReporteDepositosTotales '" & psAnioDesde & "','" & psMesDesde & "','" & psAnioHasta & "','" & psMesHasta & "','" & psAgeCod & "'"
    Set ReporteDepositosTotales = oConect.CargaRecordSet(sql)
    Exit Function
ErrReporteDepositosTotales:
        Call RaiseError(MyUnhandledError, "DCajaGeneral:ReporteDepositosTotales")
End Function
'FIN FRHU20140121 RQ13826
'PASIERS1242014
Public Function GetDatosPagoDevolucionProveedorSunat(ByVal pdFechaIni As Date, ByVal pdFechaFin As Date) As ADODB.Recordset
    Dim sql As String
    On Error GoTo ErrGetDatosPagoDevolucionProveedorSunat
    sql = "stp_sel_ERS1242014_GetDatosPagoDevolucionProveedorSUNAT '" & Format(pdFechaIni, "yyyyMMdd") & "','" & Format(pdFechaFin, "yyyyMMdd") & "'"
    Set GetDatosPagoDevolucionProveedorSunat = oConect.CargaRecordSet(sql)
    Exit Function
ErrGetDatosPagoDevolucionProveedorSunat:
    MsgBox Err.Description, vbExclamation, "Aviso"
End Function
Public Function ObtieneIFIOrigenxDevSUNAT(ByVal psIFItpo As String, ByVal psPersCodIFI As String, ByVal psCtaIFCod As String) As ADODB.Recordset
    Dim sql As String
    On Error GoTo ErrObtieneIFIOrigenxDevSUNAT
    sql = "stp_sel_ERS124_2014_ObtieneIFIOrigenxDevSUNAT '" & psIFItpo & "','" & psPersCodIFI & "','" & psCtaIFCod & "'"
    Set ObtieneIFIOrigenxDevSUNAT = oConect.CargaRecordSet(sql)
    Exit Function
ErrObtieneIFIOrigenxDevSUNAT:
    MsgBox Err.Description, vbExclamation, "Aviso"
End Function
'end PASI
