VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NContImpreReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim vsBaseComunes As String
Dim vsBasePersonas As String
Dim vsAgeNom As String
Dim vsEmpNom As String
Dim vsFecSis As String

Dim sSql    As String
Dim rs      As ADODB.Recordset
Dim dbConec As DConecta

Public Event BarraShow(pnMax)
Public Event BarraProgress(value, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
Public Event BarraClose()
'****Comentado por ALPA
Public Function GetMovCompraGastosMayor(psCtaS As String, psDocs As String, pdFechaDel As Date, pdFechaAl As Date, sCtaIGV As String, sCtaISC As String, Optional plFiltroDoc As Boolean = False, Optional pbOrdDoc As Boolean = False, _
                                        Optional sCtaRetrac As String = "", Optional sCtaRetCuarta As String = "", Optional sRUC As String = "", Optional sOpecod As String, Optional sCtaIGVND As String) As Recordset

Dim oImpuesto As DImpuesto
Set oImpuesto = New DImpuesto
Dim lsCtaCVPerdida As String
Dim oCon As NConstSistemas
Set oCon = New NConstSistemas

Dim lnIGVValor As Currency
lnIGVValor = 1 + (oImpuesto.CargaImpuestoFechaValor(oCon.LeeConstSistema(40), CDate(oCon.LeeConstSistema(15))) / 100)
lsCtaCVPerdida = oCon.LeeConstSistema(21)

Dim sObj As String, sImp As String, sDoc As String
'Se agrego el monto de igv de los proveedores no domiciliados'
' GITU 11-12-2008

    If Not pbOrdDoc Then
        
        sObj = " LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '" & sCtaIGV & "') as OtroItmIGV ON (OtroItmIGV.nMovNro = a.nMovNro and OtroItmIGV.nMovItem = a.nMovItem) " _
             & " LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '" & sCtaIGVND & "') as OtroItmIGVND ON (OtroItmIGVND.nMovNro = a.nMovNro and OtroItmIGVND.nMovItem = a.nMovItem) " _
             & " LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '" & sCtaISC & "') as OtroItmISC ON (OtroItmISC.nMovNro = a.nMovNro and OtroItmISC.nMovItem = a.nMovItem) " _
             & " LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '" & sCtaRetrac & "') as OtroItmRetrac ON (OtroItmRetrac.nMovNro = a.nMovNro and OtroItmRetrac.nMovItem = a.nMovItem)  "
        
        If Len(psDocs) > 0 Then
'            psDocs = Replace(psDocs, "'", "")
           sDoc = IIf(plFiltroDoc, "", " LEFT") & " JOIN (SELECT * FROM MovDoc   WHERE nDocTpo IN (" & psDocs & ") ) as Doc ON (Doc.nMovNro = m.nMovNro) "
        End If
        
        '***Modificado por ELRO el 20111011, según Acta 278-2011/TI-D
        'sSQL = "SELECT c.cCtaContCod, c.cCtaContDesc, ISNULL(SUM(z.nTotPV),0) as nPV, ISNULL(SUM(CASE WHEN Substring(z.cOpeCod,3,1) <> '2' THEN  z.nTotIGV WHEN SubString(z.cOpeCod,3,1) = '2' and z.nTotIGV <> 0 THEN Round(Z.nTotPV - (z.nTotPV / " & lnIGVValor & "), 2) END ),0) as nIGV, ISNULL(SUM(z.nTotISC),0) as nISC " _
        '& "FROM   " & vsBaseComunes & "CtaCont c, " _
        '& "     ( SELECT a.cCtaContCod, m.cOpeCod, ISNULL(SUM(a.nMovImporte),0) as nTotPV, SUM(ISNULL(OtroItmIGV.nMovOtroImporte,0)+ISNULL(OtroItmIGVND.nMovOtroImporte,0)) as nTotIGV, SUM(ISNULL(OtroItmISC.nMovOtroImporte,0)) as nTotISC" _
        '& "       FROM   Mov m JOIN MovCta a ON m.nMovNro = a.nMovNro " & sObj & sImp & sDoc & " " _
        '& "       WHERE  m.nMovEstado = '" & gMovEstContabMovContable & "' and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & "," & gMovFlagExtornado & ") and a.cCtaContCod IN (" & psCtaS & ") and " _
        '& "             substring(m.cMovNro,1,8) BETWEEN '" & Format(pdFechaDel, "yyyymmdd") & "' and '" & Format(pdFechaAl, "yyyymmdd") & "' " _
        '& "       GROUP BY a.cCtaContCod, m.cOpeCod " _
        '& "     ) as z " _
        '& "WHERE  z.cCtaContCod LIKE c.cCtaContCod + '%' " _
        '& "GROUP BY c.cCtaContCod, c.cCtaContDesc ORDER BY c.cCtaContCod"
        
        sSql = "SELECT c.cCtaContCod," _
                    & " c.cCtaContDesc," _
                    & " round(ISNULL(SUM(z.nTotPV),0),2) as nPV," _
                    & " round(ISNULL(SUM(CASE WHEN Substring(z.cOpeCod,3,1) <> '2' THEN  z.nTotIGV WHEN SubString(z.cOpeCod,3,1) = '2' and z.nTotIGV <> 0 THEN Round(Z.nTotPV - (z.nTotPV / " & lnIGVValor & "), 2) END ),0),2) as nIGV," _
                    & " round(ISNULL(SUM(z.nTotISC),0),2) as nISC," _
                    & "(SELECT ROUND(ISNULL(SUM(CS.nCtaSaldoImporte),0),2) " _
                    & " FROM CtaSaldo CS WHERE   CS.cCtaContCod  LIKE c.cCtaContCod + '%' " _
                    & " AND CS.dCtaSaldoFecha = (SELECT MAX(CS1.dCtaSaldoFecha) FROM CtaSaldo CS1 " _
                    & " WHERE CS1.cCtaContCod = CS.cCtaContCod and CS1.dCtaSaldoFecha <= '" & Format(pdFechaDel - 1, gsFormatoFecha) & "') ) as Acumulado" _
                    & " FROM   " & vsBaseComunes & "CtaCont c, " _
                    & "     ( SELECT a.cCtaContCod, m.cOpeCod, ISNULL(SUM(a.nMovImporte),0) as nTotPV, SUM(ISNULL(OtroItmIGV.nMovOtroImporte,0)+ISNULL(OtroItmIGVND.nMovOtroImporte,0)) as nTotIGV, SUM(ISNULL(OtroItmISC.nMovOtroImporte,0)) as nTotISC" _
                    & "       FROM   Mov m JOIN MovCta a ON m.nMovNro = a.nMovNro " & sObj & sImp & sDoc & " " _
                    & "       WHERE  m.nMovEstado = '" & gMovEstContabMovContable & "' and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & "," & gMovFlagExtornado & ") and a.cCtaContCod IN (" & psCtaS & ") and " _
                    & "             substring(m.cMovNro,1,8) BETWEEN '" & Format(pdFechaDel, "yyyymmdd") & "' and '" & Format(pdFechaAl, "yyyymmdd") & "' " _
                    & "       GROUP BY a.cCtaContCod, m.cOpeCod " _
                    & "     ) as z " _
                    & "WHERE  z.cCtaContCod LIKE c.cCtaContCod + '%' " _
                    & "GROUP BY c.cCtaContCod, c.cCtaContDesc ORDER BY c.cCtaContCod"
        '***Fin Modificado por ELRO**********************************
    Else
    
'*** PEAC 20110308
'        sObj = " " _
'             & " LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '" & sCtaIGV & "') as OtroItmIGV ON (OtroItmIGV.nMovNro = a.nMovNro and OtroItmIGV.nMovItem = a.nMovItem) " _
'             & " LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '" & sCtaIGVND & "') as OtroItmIGVND ON (OtroItmIGVND.nMovNro = a.nMovNro and OtroItmIGVND.nMovItem = a.nMovItem) " _
'             & " LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '" & sCtaISC & "') as OtroItmISC ON (OtroItmISC.nMovNro = a.nMovNro and OtroItmISC.nMovItem = a.nMovItem) " _
'             & " LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '" & sCtaRetrac & "') as OtroItmRetrac ON (OtroItmRetrac.nMovNro = a.nMovNro and OtroItmRetrac.nMovItem = a.nMovItem)  " _
'             & " LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '" & sCtaRetCuarta & "') as OtroItmRentaCuarta ON (OtroItmRentaCuarta.nMovNro = a.nMovNro and OtroItmRentaCuarta.nMovItem = a.nMovItem)  "

        If Len(Trim(psDocs)) > 0 Then
            psDocs = Replace(psDocs, "'", "")
'*** PEAC 20110308
'           sDoc = IIf(plFiltroDoc, "", " INNER") & " JOIN (SELECT * FROM MovDoc   WHERE nDocTpo IN (" & psDocs & ") ) as Doc ON (Doc.nMovNro = m.nMovNro) "
        End If
'ALPA 20090304***************************************************************************
'        sSQL = "SELECT  isnull(z.nMovTpoCambio,0)nMovTpoCambio,OFIS=SUBSTRING(z.cMovNro,18,2), z.cDocNro, z.cMovNro, z.cMovDesc, z.nDocTpo, z.cDocDesc, z.dDocFecha, z.nMovNro,Tipo=z.cDocAbrev, z.cOpeCod, ISNULL(SUM(z.nTotPVME),0)  - ISNULL(SUM(z.nTotIGV),0) as nPVME , ISNULL(SUM(z.nTotPV),0) as nPV, ISNULL(SUM(CASE WHEN Substring(z.cOpeCod,3,1) <> '2' THEN  z.nTotIGV WHEN SubString(z.cOpeCod,3,1) = '2' and z.nTotIGV <> 0 THEN Round(Z.nTotPV - (z.nTotPV / " & lnIGVValor & "), 2) END ),0) as nIGV, ISNULL(SUM(z.nTotISC),0) as nISC,ISNULL(SUM(z.nTotRetrac),0)AS nRetracc," _
'             & " ISNULL(SUM(z.nTotRentaCuarta),0)AS nRentaCuarta ," _
'             & " IsNull((Select PE.cPersNombre  FROM MOVGASTO MG Inner Join Persona PE ON MG.cPersCod = PE.cPersCod Where MG.nMoVnro = z.nMovNro),'') cPersNombre, IsNull((Select PE.cPersIDnro FROM MOVGASTO MG  Inner Join PersID PE ON MG.cPersCod = PE.cPersCod And cPersIDTpo = 2 Where MG.nMoVnro = z.nMovNro),isnull((Select PE.cPersIDnro FROM MOVGASTO MG Inner Join PersID PE ON MG.cPersCod = PE.cPersCod And cPersIDTpo = 1  Where MG.nMoVnro = z.nMovNro),isnull((Select PE.cPersIDnro FROM MOVGASTO MG Inner Join PersID PE ON MG.cPersCod = PE.cPersCod And cPersIDTpo = 4  Where MG.nMoVnro = z.nMovNro),''))) cRUC," _
'             & " IsNull((Select Top 1  MD.nDocTpo From MovRef MR Inner Join MovDoc MD On MR.nMovNroref = MD.nMovNro And MD.nDocTpo = 33 Where MR.nMovNro = z.nMovnro),'') cOC ," _
'             & " (Select MDO.dFecha From MovDocOtros MDO inner Join Mov M On M.nMovNro = MDO.nMovNro Where MDO.nMovNro = z.nMovnro) FechaRef, " _
'             & " IsNull((Select PE.cPersNombre  FROM MOVGASTOOtros MGO Inner Join Persona PE ON MGO.cPersCod = PE.cPersCod Where MGO.nMoVnro = z.nMovNro),'') cPersNombreRef " _
'             & "  FROM ( SELECT Doc.cDocNro, m.nMovNro, MTC.nMovTpoCambio , m.cOpeCod , Doc.nDocTpo, Doc.dDocFecha  , m.cMovNro, m.cMovDesc,Docum.cDocAbrev,Docum.cDocDesc, ISNULL(SUM(a.nMovImporte),0) as nTotPV, SUM(ISNULL(OtroItmIGV.nMovOtroImporte,0)+ISNULL(OtroItmIGVND.nMovOtroImporte,0)) as nTotIGV, SUM(ISNULL(OtroItmISC.nMovOtroImporte,0)) as nTotISC, SUM(ISNULL(OtroItmRetrac.nMovOtroImporte,0)) as nTotRetrac," _
'             & "              Case When SubString(m.cOpeCod,3,1) = '2' And MTC.nMovTpoCambio Is Not Null Then SUM(ISNULL(OtroItmRentaCuarta.nMovOtroImporte,0))* MTC.nMovTpoCambio Else SUM(ISNULL(OtroItmRentaCuarta.nMovOtroImporte,0)) End as nTotRentaCuarta ,ISNULL(SUM(ME.nMovMEImporte),0) as nTotPVME  " _
'             & "       FROM   Mov m JOIN MovCta a ON m.nMovNro = a.nMovNro " & sObj & sImp & sDoc & " INNER JOIN Documento Docum ON Docum.nDocTpo=Doc.nDocTpo " _
'             & "       Inner Join Movgasto MG ON m.nMovnro = MG.nMovnro " _
'             & "       Left Join MovME ME On m.nMovNro = ME.nMovNro and a.nMovItem = me.nMovItem " _
'             & "       Left Join MovTpoCambio MTC On m.nMovNro = MTC.nMovNro " _
'             & "       WHERE  m.nMovEstado = '" & gMovEstContabMovContable & "' and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") AND A.nMovImporte > 0 "

'*** PEAC 20110308
'        sSQL = "SELECT  isnull(z.nMovTpoCambio,0)nMovTpoCambio,OFIS=SUBSTRING(z.cMovNro,18,2), z.cDocNro, z.cMovNro, z.cMovDesc, z.nDocTpo, z.cDocDesc, z.dDocFecha, z.nMovNro,Tipo=z.cDocAbrev, z.cOpeCod, ISNULL(SUM(z.nTotPVME),0)  - ISNULL(SUM(z.nTotIGV),0) as nPVME , ISNULL(SUM(z.nTotPV),0) as nPV, ISNULL(SUM(CASE WHEN Substring(z.cOpeCod,3,1) <> '2' THEN  z.nTotIGV WHEN SubString(z.cOpeCod,3,1) = '2' and z.nTotIGV <> 0 THEN Round(Z.nTotPV - (z.nTotPV / " & lnIGVValor & "), 2) END ),0) as nIGV, ISNULL(SUM(z.nTotISC),0) as nISC,ISNULL(SUM(z.nTotRetrac),0)AS nRetracc,"
'        sSQL = sSQL & " ISNULL(SUM(z.nTotRentaCuarta),0)AS nRentaCuarta ,"
'        sSQL = sSQL & " IsNull((Select PE.cPersNombre  FROM MOVGASTO MG Inner Join Persona PE ON MG.cPersCod = PE.cPersCod Where MG.nMoVnro = z.nMovNro),'') cPersNombre, IsNull((Select PE.cPersIDnro FROM MOVGASTO MG  Inner Join PersID PE ON MG.cPersCod = PE.cPersCod And cPersIDTpo = 2 Where MG.nMoVnro = z.nMovNro),isnull((Select PE.cPersIDnro FROM MOVGASTO MG Inner Join PersID PE ON MG.cPersCod = PE.cPersCod And cPersIDTpo = 1  Where MG.nMoVnro = z.nMovNro),isnull((Select PE.cPersIDnro FROM MOVGASTO MG Inner Join PersID PE ON MG.cPersCod = PE.cPersCod And cPersIDTpo = 4  Where MG.nMoVnro = z.nMovNro),''))) cRUC,"
'        sSQL = sSQL & " IsNull((Select Top 1  MD.nDocTpo From MovRef MR Inner Join MovDoc MD On MR.nMovNroref = MD.nMovNro And MD.nDocTpo = 33 Where MR.nMovNro = z.nMovnro),'') cOC ,"
'        sSQL = sSQL & " (Select MDO.dFecha From MovDocOtros MDO inner Join Mov M On M.nMovNro = MDO.nMovNro Where MDO.nMovNro = z.nMovnro) FechaRef, "
'        sSQL = sSQL & " IsNull((Select PE.cPersNombre  FROM MOVGASTOOtros MGO Inner Join Persona PE ON MGO.cPersCod = PE.cPersCod Where MGO.nMoVnro = z.nMovNro),'') cPersNombreRef, "
'        sSQL = sSQL & " IsNull((Select PE.cPersIDnro FROM MOVGASTOOTROS MG  Inner Join PersID PE ON MG.cPersCod = PE.cPersCod And cPersIDTpo = 2 Where MG.nMoVnro = z.nMovNro),isnull((Select PE.cPersIDnro FROM MOVGASTOOTROS MG Inner Join PersID PE ON MG.cPersCod = PE.cPersCod And cPersIDTpo = 1  Where MG.nMoVnro = z.nMovNro),isnull((Select PE.cPersIDnro FROM MOVGASTOOTROS MG Inner Join PersID PE ON MG.cPersCod = PE.cPersCod And cPersIDTpo = 4  Where MG.nMoVnro = z.nMovNro),''))) cRUCRef "
'        sSQL = sSQL & "  FROM ( SELECT Doc.cDocNro, m.nMovNro, MTC.nMovTpoCambio , m.cOpeCod , Doc.nDocTpo, Doc.dDocFecha  , m.cMovNro, m.cMovDesc,Docum.cDocAbrev,Docum.cDocDesc, ISNULL(SUM(a.nMovImporte),0) as nTotPV, SUM(ISNULL(OtroItmIGV.nMovOtroImporte,0)+ISNULL(OtroItmIGVND.nMovOtroImporte,0)) as nTotIGV, SUM(ISNULL(OtroItmISC.nMovOtroImporte,0)) as nTotISC, SUM(ISNULL(OtroItmRetrac.nMovOtroImporte,0)) as nTotRetrac,"
'        sSQL = sSQL & "              Case When SubString(m.cOpeCod,3,1) = '2' And MTC.nMovTpoCambio Is Not Null Then SUM(ISNULL(OtroItmRentaCuarta.nMovOtroImporte,0))* MTC.nMovTpoCambio Else SUM(ISNULL(OtroItmRentaCuarta.nMovOtroImporte,0)) End as nTotRentaCuarta ,ISNULL(SUM(ME.nMovMEImporte),0) as nTotPVME  "
'        sSQL = sSQL & "       FROM   Mov m JOIN MovCta a ON m.nMovNro = a.nMovNro " & sObj & sImp & sDoc & " INNER JOIN Documento Docum ON Docum.nDocTpo=Doc.nDocTpo "
'        sSQL = sSQL & "       Inner Join Movgasto MG ON m.nMovnro = MG.nMovnro "
'        sSQL = sSQL & "       Left Join MovME ME On m.nMovNro = ME.nMovNro and a.nMovItem = me.nMovItem "
'        sSQL = sSQL & "       Left Join MovTpoCambio MTC On m.nMovNro = MTC.nMovNro "
'        sSQL = sSQL & "       WHERE  m.nMovEstado = '" & gMovEstContabMovContable & "' and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") AND A.nMovImporte > 0 "
'
'****************************************************************************************************
        '-- Se agrego la condición por el codigo de operacion porque el reporte --'
        '-- de analisis de gasto utiliza el mismo query GITU 17/04/2008 --'
        
        If sOpecod = "760200" Then
            psCtaS = Replace(psCtaS, "'", "")
'            sSQL = sSQL & " and a.cCtaContCod IN (" & psCtaS & ") " '*** PEAC 20110308
        End If

'*** PEAC 20110308
'        sSQL = sSQL & "        and substring(m.cMovNro,1,8) BETWEEN '" & Format(pdFechaDel, "yyyymmdd") & "' and '" & Format(pdFechaAl, "yyyymmdd") & "' and a.cctacontcod NOT Like '" & Trim(lsCtaCVPerdida) & "%' " _
             & "       GROUP BY Doc.cDocNro, m.cMovNro,MTC.nMovTpoCambio, m.cMovDesc, m.cOpeCod, m.nMovNro, Doc.nDocTpo, Doc.dDocFecha, Docum.cDocAbrev ,Docum.cDocDesc " _
             & "     ) as z "
        '-- Fin GITU
        
        
'*** PEAC 20110308
'             '****Comentado por ALPA
'             If sRUC <> "" Then
'
'             sSQL = sSQL & "          where IsNull((Select PE.cPersIDnro" _
'                & "      FROM MOVGASTO MG  Inner Join PersID PE ON MG.cPersCod = PE.cPersCod And cPersIDTpo = 2 " _
'                & "      Where MG.nMoVnro = z.nMovNro)," _
'                & "      isnull((Select PE.cPersIDnro FROM MOVGASTO MG Inner Join PersID PE ON MG.cPersCod = PE.cPersCod " _
'                & "      And cPersIDTpo = 1  Where MG.nMoVnro = z.nMovNro)," _
'                & "     isnull((Select PE.cPersIDnro " _
'                & "      FROM MOVGASTO MG Inner Join PersID PE " _
'                & "     ON MG.cPersCod = PE.cPersCod And cPersIDTpo = 4 " _
'                & "      Where MG.nMoVnro = z.nMovNro),'')))='" + sRUC + "' "
'            '****
'           End If
           
'*** PEAC 20110308
'             sSQL = sSQL & " GROUP BY z.dDocFecha, z.cDocNro, z.cMovNro, z.cMovDesc, z.cOpeCod, z.nMovNro, z.nDocTpo, z.nMovTpoCambio,z.cDocAbrev,z.cDocDesc "
             
'***Comentado por ALPA*****************************************************************************************************************************
'***28/03/2008*************************************************************************************************************************************
             
'*** PEAC 20110308
'             sSQL = sSQL & " ORDER BY  z.nDocTpo,z.dDocFecha"
             
'***End********************************************************************************************************************************************

    '*** PEAC 20110308
    sSql = "exec stp_sel_ObtieneMovCompraGastosMayor '" & Format(pdFechaDel, "yyyymmdd") & "','" & Format(pdFechaAl, "yyyymmdd") & "','" & sOpecod & "','" & psCtaS & "','" & Trim(lsCtaCVPerdida) & "','" & psDocs & "','" & Trim(sRUC) & "'"

    End If
    '***Comentado por ALPA
    '***Direncias de Union en
'    sSQL = sSQL & "Union All " _
'         & " SELECT  isnull(z.nMovTpoCambio,0)nMovTpoCambio,OFIS=SUBSTRING(z.cMovNro,18,2), z.cDocNro, z.cMovNro, z.cMovDesc, z.nDocTpo, z.cDocDesc, z.dDocFecha, z.nMovNro, " _
'         & " Tipo=z.cDocAbrev, z.cOpeCod, ISNULL(SUM(z.nTotPVME),0)  - ISNULL(SUM(z.nTotIGV),0) as nPVME , " _
'         & " ISNULL((SUM(z.nTotPV)*-1),0) as nPV, ISNULL(SUM(CASE WHEN Substring(z.cOpeCod,3,1) <> '2' THEN  (z.nTotIGV*-1) WHEN SubString(z.cOpeCod,3,1) = '2' and z.nTotIGV <> 0 THEN (Round(Z.nTotPV - (z.nTotPV / 1.19), 2)*-1) END ),0) as nIGV, " _
'         & " ISNULL(SUM(z.nTotISC),0) as nISC,ISNULL(SUM(z.nTotRetrac),0)AS nRetracc,ISNULL(SUM(z.nTotRentaCuarta),0)AS nRentaCuarta , " _
'         & " IsNull((Select PE.cPersNombre  FROM MOVGASTO MG Inner Join Persona PE ON MG.cPersCod = PE.cPersCod Where MG.nMoVnro = z.nMovNro),'') cPersNombre, " _
'         & " IsNull((Select PE.cPersIDnro FROM MOVGASTO MG  Inner Join PersID PE ON MG.cPersCod = PE.cPersCod And cPersIDTpo = 2 Where MG.nMoVnro = z.nMovNro),isnull((Select PE.cPersIDnro FROM MOVGASTO MG Inner Join PersID PE ON MG.cPersCod = PE.cPersCod And cPersIDTpo = 1  Where MG.nMoVnro = z.nMovNro),isnull((Select PE.cPersIDnro FROM MOVGASTO MG Inner Join PersID PE ON MG.cPersCod = PE.cPersCod And cPersIDTpo = 4  Where MG.nMoVnro = z.nMovNro),''))) cRUC, " _
'         & " IsNull((Select Top 1  MD.nDocTpo From MovRef MR Inner Join MovDoc MD On MR.nMovNroref = MD.nMovNro And MD.nDocTpo = 33 Where MR.nMovNro = z.nMovnro),'') cOC , " _
'         & " (Select MDO.dFecha From MovDocOtros MDO inner Join Mov M On M.nMovNro = MDO.nMovNro Where MDO.nMovNro = z.nMovnro) FechaRef, " _
'         & " IsNull((Select PE.cPersNombre  FROM MOVGASTOOtros MGO Inner Join Persona PE ON MGO.cPersCod = PE.cPersCod Where MGO.nMoVnro = z.nMovNro),'') cPersNombreRef " _
'         & " FROM ( SELECT Doc.cDocNro, m.nMovNro, MTC.nMovTpoCambio , m.cOpeCod , Doc.nDocTpo, " _
'         & "              Doc.dDocFecha  , m.cMovNro, m.cMovDesc,Docum.cDocAbrev,Docum.cDocDesc, " _
'         & " ISNULL(SUM(a.nMovImporte),0) as nTotPV, SUM(ISNULL(OtroItmIGV.nMovOtroImporte,0)) as nTotIGV," _
'         & " SUM(ISNULL(OtroItmISC.nMovOtroImporte,0)) as nTotISC, SUM(ISNULL(OtroItmRetrac.nMovOtroImporte,0)) as nTotRetrac, " _
'         & " SUM(ISNULL(OtroItmRentaCuarta.nMovOtroImporte,0)) as nTotRentaCuarta , ISNULL(SUM(Me.nMovMEImporte), 0) As nTotPVME " _
'         & " FROM   Mov m  JOIN MovCta a ON m.nMovNro = a.nMovNro " _
'         & "       LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '21140102') as OtroItmIGV ON (OtroItmIGV.nMovNro = a.nMovNro and OtroItmIGV.nMovItem = a.nMovItem) " _
'         & "       LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '2918070698') as OtroItmISC ON (OtroItmISC.nMovNro = a.nMovNro and OtroItmISC.nMovItem = a.nMovItem) " _
'         & "       LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '25141912') as OtroItmRetrac ON (OtroItmRetrac.nMovNro = a.nMovNro and OtroItmRetrac.nMovItem = a.nMovItem) " _
'         & "       LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '2114020101') as OtroItmRentaCuarta ON (OtroItmRentaCuarta.nMovNro = a.nMovNro and OtroItmRentaCuarta.nMovItem = a.nMovItem) " _
'         & "       INNER JOIN (SELECT * FROM MovDoc  WHERE nDocTpo IN ('1','2','3','5','6','7','8','10','11','12','13','14','16','19','23','28','38','39','44','50','80','93','95','96','128','131','140') ) as Doc ON (Doc.nMovNro = m.nMovNro) " _
'         & "       INNER JOIN Documento Docum ON Docum.nDocTpo=Doc.nDocTpo" _
'         & "       Left Join MovME ME On m.nMovNro = ME.nMovNro and a.nMovItem = me.nMovItem " _
'         & "       Left Join MovTpoCambio MTC On m.nMovNro = MTC.nMovNro " _
'         & " WHERE  m.nMovEstado = '" & gMovEstContabMovContable & "' and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ")  AND A.nMovImporte < 0 "
'    sSQL = sSQL & "and a.cCtaContCod IN ('451301070110') " _
'         & " and substring(m.cMovNro,1,8) BETWEEN '" & Format(pdFechaDel, "yyyymmdd") & "' and '" & Format(pdFechaAl, "yyyymmdd") & "' and a.cctacontcod NOT Like '" & Trim(lsCtaCVPerdida) & "%' " _
'         & " GROUP BY Doc.cDocNro, m.cMovNro,MTC.nMovTpoCambio, m.cMovDesc, m.cOpeCod, m.nMovNro, Doc.nDocTpo, Doc.dDocFecha, Docum.cDocAbrev ,Docum.cDocDesc      ) as z " _
'         & " GROUP BY z.dDocFecha, z.cDocNro, z.cMovNro, z.cMovDesc, z.cOpeCod, z.nMovNro, z.nDocTpo, z.nMovTpoCambio,z.cDocAbrev,z.cDocDesc" _
'         & " ORDER BY  z.nDocTpo,z.dDocFecha "
   '  End If
      
      
'    psDocs = Replace(psDocs, "'", "")
   
'   sSQL = "exec stp_sel_MovCompraGastosMayor '" & sCtaIGV & "','" & sCtaIGVND & "','" & _
'            sCtaISC & "','" & sCtaRetrac & "','" & sCtaRetCuarta & "','" & psDocs & "'," & _
'            lnIGVValor & ",'" & sOpecod & "','" & _
'            psCtaS & "','" & Format(pdFechaDel, "yyyymmdd") & "','" & _
'            Format(pdFechaAl, "yyyymmdd") & "','" & Trim(lsCtaCVPerdida) & "','" & _
'            sRUC & "'"

    Set GetMovCompraGastosMayor = dbConec.CargaRecordSet(sSql)
End Function

'Public Function GetMovCompraGastos(sCtas As String, sDocs As String, pdFechaDel As Date, pdFechaAl As Date, psCtaIgv As String, psCtaISC As String, Optional plFiltroDoc As Boolean = False, Optional pbOrdenDoc As Boolean = False) As Recordset
'Dim sObj As String, sDoc As String, sImp As String
'
'sObj = " LEFT JOIN (SELECT mg.nMovNro, per.cPersNombre, ISNULL(pid.cPersIDnro,'') cRuc FROM MovGasto mg JOIN " & vsBasePersonas & "Persona Per ON Per.cPersCod = mg.cPersCod LEFT JOIN ( SELECT * FROM " & vsBasePersonas & "PersID WHERE cPersIDTpo = '" & gPersIdRUC & "' ) pid ON pid.cPersCod = per.cPersCod " _
'     & "           ) as Obj ON Obj.nMovNro = a.nMovNro " _
'     & " LEFT JOIN (SELECT nMovNro, nMovItem, Sum(nMovOtroImporte) as nMovOtroImporte  FROM MovOtrosItem WHERE cMovOtroVariable = '" & psCtaIgv & "' GROUP BY nMovNro, nMovItem) as OtroItmIGV ON OtroItmIGV.nMovNro = a.nMovNro and OtroItmIGV.nMovItem = a.nMovItem " _
'     & " LEFT JOIN (SELECT nMovNro, nMovItem, Sum(nMovOtroImporte) as nMovOtroImporte  FROM MovOtrosItem WHERE cMovOtroVariable = '" & psCtaISC & "' GROUP BY nMovNro, nMovItem) as OtroItmISC ON OtroItmISC.nMovNro = a.nMovNro and OtroItmISC.nMovItem = a.nMovItem "
'
'If sDocs <> "" Then
'   sDoc = IIf(plFiltroDoc, "", " LEFT") & " JOIN (SELECT * FROM MovDoc   WHERE nDocTpo IN (" & sDocs & ") ) as Doc ON (Doc.nMovNro = b.nMovNro) "
'End If
'
'sSQL = "SELECT a.cCtaContCod, c.cCtaContDesc, b.cMovNro, b.cOpeCod, Doc.nDocTpo, Doc.cDocNro, " _
'     & "       Doc.dDocFecha, b.cMovDesc, Obj.cPersNombre, Obj.cRuc, ISNULL(SUM(a.nMovImporte),0) as nPV, " _
'     & "       ISNULL(SUM(OtroItmIGV.nMovOtroImporte),0) as nIGV, " _
'     & "       ISNULL(SUM(OtroItmISC.nMovOtroImporte),0) as nISC, ISNULL(mdChq.cDocNro,'') cDocNroEmi, ISNULL( mdChq.dDocFecha, '') dDocFechaEmi " _
'     & "FROM   Mov b JOIN MovCta a ON a.nMovNro = b.nMovNro " & sObj & sDoc & sImp _
'     & "       LEFT JOIN MovRef mr ON mr.nMovNro = b.nMovNro " _
'     & "       LEFT JOIN MOVDOC mdChq ON mdChq.nMovNro = mr.nMovNroRef and mdChq.nDocTpo IN (" & TpoDocCheque & " ) " _
'     & ", " & vsBaseComunes & "CtaCont c " _
'     & "WHERE  b.nMovEstado = '" & gMovEstContabMovContable & "' and not b.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & ") and a.cCtaContCod IN (" & sCtas & ") and " _
'     & "       c.cCtaContCod = a.cCtaContCod and " _
'     & "       substring(b.cMovNro,1,8) BETWEEN '" & Format(pdFechaDel, "yyyymmdd") & "' and '" & Format(pdFechaAl, "yyyymmdd") & "' " _
'     & "GROUP BY a.cCtaContCod, c.cCtaContDesc, b.cMovNro, b.cOpeCod, Doc.nDocTpo, Doc.cDocNro, Doc.dDocFecha, " _
'     & "       b.cMovDesc, Obj.cPersNombre, Obj.cRuc, ISNULL(mdChq.cDocNro,''), ISNULL( mdChq.dDocFecha, '') " _
'     & "ORDER BY " & IIf(Not pbOrdenDoc, " a.cCtaContCod, b.cMovNro", " Doc.cDocNro")
'Set GetMovCompraGastos = dbConec.CargaRecordSet(sSQL)
'End Function

Public Function GetMovCompraGastos(sCtas As String, sDocs As String, pdFechaDel As Date, pdFechaAl As Date, psCtaIgv As String, psCtaISC As String, Optional plFiltroDoc As Boolean = False, Optional pbOrdenDoc As Boolean = False, Optional pbResumen As Integer) As Recordset
Dim sObj As String, sDoc As String, sImp As String
Dim lsOrden As String

sObj = " LEFT JOIN (SELECT mg.nMovNro, per.cPersNombre, ISNULL(pid.cPersIDnro,'') cRuc FROM MovGasto mg JOIN " & vsBasePersonas & "Persona Per ON Per.cPersCod = mg.cPersCod LEFT JOIN ( SELECT * FROM " & vsBasePersonas & "PersID WHERE cPersIDTpo = '" & gPersIdRUC & "' ) pid ON pid.cPersCod = per.cPersCod " _
     & "           ) as Obj ON Obj.nMovNro = a.nMovNro " _
     & " LEFT JOIN (SELECT nMovNro, nMovItem, Sum(nMovOtroImporte) as nMovOtroImporte  FROM MovOtrosItem WHERE cMovOtroVariable = '" & psCtaIgv & "' GROUP BY nMovNro, nMovItem) as OtroItmIGV ON OtroItmIGV.nMovNro = a.nMovNro and OtroItmIGV.nMovItem = a.nMovItem " _
     & " LEFT JOIN (SELECT nMovNro, nMovItem, Sum(nMovOtroImporte) as nMovOtroImporte  FROM MovOtrosItem WHERE cMovOtroVariable = '" & psCtaISC & "' GROUP BY nMovNro, nMovItem) as OtroItmISC ON OtroItmISC.nMovNro = a.nMovNro and OtroItmISC.nMovItem = a.nMovItem "
     
If sDocs <> "" Then
   sDoc = IIf(plFiltroDoc, "", " LEFT") & " JOIN (SELECT * FROM MovDoc   WHERE nDocTpo IN (" & sDocs & ") ) as Doc ON (Doc.nMovNro = b.nMovNro) "
End If
If pbResumen = 1 Then
   lsOrden = " ORDER BY Doc.nDocTpo,Doc.cDocNro " 'ordenar por tipo de Documento
Else
   lsOrden = "ORDER BY " & IIf(Not pbOrdenDoc, " a.cCtaContCod, b.cMovNro", " Doc.cDocNro")
End If
'***Comentado por ELRO el 20111128, según Acta 278-2011/TI-D
'sSql = "SELECT a.cCtaContCod, c.cCtaContDesc, b.cMovNro, b.cOpeCod, Doc.nDocTpo, Doc.cDocNro, " _
'     & "       Doc.dDocFecha, b.cMovDesc, Obj.cPersNombre, Obj.cRuc, ISNULL(SUM(a.nMovImporte),0) as nPV, " _
'     & "       ISNULL(SUM(OtroItmIGV.nMovOtroImporte),0) as nIGV, " _
'     & "       ISNULL(SUM(OtroItmISC.nMovOtroImporte),0) as nISC, ISNULL(mdChq.cDocNro,'') cDocNroEmi, ISNULL( mdChq.dDocFecha, '') dDocFechaEmi " _
'     & "FROM   Mov b JOIN MovCta a ON a.nMovNro = b.nMovNro " & sObj & sDoc & sImp _
'     & "       LEFT JOIN MovRef mr ON mr.nMovNro = b.nMovNro " _
'     & "       LEFT JOIN MOVDOC mdChq ON mdChq.nMovNro = mr.nMovNroRef and mdChq.nDocTpo IN (" & TpoDocCheque & " ) " _
'     & ", " & vsBaseComunes & "CtaCont c " _
'     & "WHERE  b.nMovEstado = '" & gMovEstContabMovContable & "' and not b.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & "," & gMovFlagExtornado & ") and a.cCtaContCod IN (" & sCtas & ") and " _
'     & "       c.cCtaContCod = a.cCtaContCod and " _
'     & "       substring(b.cMovNro,1,8) BETWEEN '" & Format(pdFechaDel, "yyyymmdd") & "' and '" & Format(pdFechaAl, "yyyymmdd") & "' " _
'     & "GROUP BY a.cCtaContCod, c.cCtaContDesc, b.cMovNro, b.cOpeCod, Doc.nDocTpo, Doc.cDocNro, Doc.dDocFecha, " _
'     & "       b.cMovDesc, Obj.cPersNombre, Obj.cRuc, ISNULL(mdChq.cDocNro,''), " _
'     & " ISNULL( mdChq.dDocFecha, '')" & lsOrden
'Fin Comentado por ELRO
     '& "ORDER BY " & IIf(Not pbOrdenDoc, " a.cCtaContCod, b.cMovNro", " Doc.cDocNro")
'***Modificado por ELRO el 20111128, según Acta 278-2011/TI-D
sSql = "SELECT a.cCtaContCod, c.cCtaContDesc, b.cMovNro, b.cOpeCod, Doc.nDocTpo, Doc.cDocNro, " _
     & "       Doc.dDocFecha, b.cMovDesc, Obj.cPersNombre, Obj.cRuc, ISNULL(SUM(a.nMovImporte),0) as nPV, " _
     & "       ISNULL(SUM(OtroItmIGV.nMovOtroImporte),0) as nIGV, " _
     & "       ISNULL(SUM(OtroItmISC.nMovOtroImporte),0) as nISC, ISNULL(mdChq.cDocNro,'') cDocNroEmi, ISNULL( mdChq.dDocFecha, '') dDocFechaEmi " _
     & "FROM   Mov b JOIN MovCta a ON a.nMovNro = b.nMovNro " & sObj & sDoc & sImp _
     & "       LEFT JOIN MOVDOC mdChq ON mdChq.nMovNro = b.nMovNro and mdChq.nDocTpo IN (" & TpoDocCheque & " ) " _
     & ", " & vsBaseComunes & "CtaCont c " _
     & "WHERE  b.nMovEstado = '" & gMovEstContabMovContable & "' and not b.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & "," & gMovFlagExtornado & ") and a.cCtaContCod IN (" & sCtas & ") and " _
     & "       c.cCtaContCod = a.cCtaContCod and " _
     & "       substring(b.cMovNro,1,8) BETWEEN '" & Format(pdFechaDel, "yyyymmdd") & "' and '" & Format(pdFechaAl, "yyyymmdd") & "' " _
     & "GROUP BY a.cCtaContCod, c.cCtaContDesc, b.cMovNro, b.cOpeCod, Doc.nDocTpo, Doc.cDocNro, Doc.dDocFecha, " _
     & "       b.cMovDesc, Obj.cPersNombre, Obj.cRuc, ISNULL(mdChq.cDocNro,''), " _
     & " ISNULL( mdChq.dDocFecha, '')" & lsOrden
'Fin Modificado por ELRO
Set GetMovCompraGastos = dbConec.CargaRecordSet(sSql)
End Function

'Private Function ImprimeRegCompraGastosCab(ByRef nLin As Integer, ByRef P As Integer) As String
'Dim K As Integer, nObjs As Integer
'Dim sTexto As String
'If nLin > gnLinPage - 6 Then
'   If P > 0 Then sTexto = oImpresora.gPrnSaltoPagina
'   P = P + 1
'   Linea sTexto, COFF & " CMAC - TRUJILLO " & Space(42) & gdFecSis & " - " & Format(Time, "hh:mm:ss")
'   Linea sTexto, Space(72) & "Pag. " & Format(P, "000")
'   Linea sTexto, BON & Centra(" R E G I S T R O   D E   C O M P R A S  /  G A S T O S ", gnColPage)
'   Linea sTexto, Centra("( DEL " & dFecha & " AL " & dFecha2 & " )", gnColPage) & BOFF
'   Linea sTexto, CON, 0
'   Linea sTexto, " ========================================================================================================================================================================================================================"
'   Linea sTexto, "  Fecha                 D O C U M E N T O                       P R O V E E D O R                                                                                    VALOR                                       PRECIO  "
'   Linea sTexto, "  de                   ------------------------------------ ------------------------------------------------------      C O N C E P T O                              VENTA         I.G.V.         I.S.C.          VENTA  "
'   Linea sTexto, "  Registro              Fecha    Tpo  Número                 R.U.C.    Apellidos y Nombres / Razón Social                                                                                                                "
'   Linea sTexto, " ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------" & COFF
'   nLin = 9
'   ImprimeRegCompraGastosCab = sTexto
'End If
'End Function

Public Function CargaCOAProveedores(Optional dFecha As String = "000000") As Recordset
''sSQL = "SELECT p.cPersCod, p.cPersNombre cNomPers, pid.cPersIDNro cProvRuc "
''      & "FROM Persona p join PersID pid ON pid.cPersCod = p.cPersCod "
''      & "WHERE pid.cPersIDTpo = '" & gPersIdRUC & "' "
''      & "      and cPersIDNro is not NULL and cPersIDNro <> '' "
''      & "      and cPersIDNro <> '00000000' and LEN(cPersIDNro) = 11 "

'sSQL = "SELECT P.cPerscod,ISNULL(P.cPersNombre,'') cNomPers,ISNULL(pid.cPersIDnro,'') as cProvRuc  " & _
'       "FROM MOV M " & _
'       "JOIN MOVCTA MC ON M.nMovNro = MC.nMovNro " & _
'       "LEFT JOIN MovME me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " & _
'       "JOIN MOVDOC MD ON M.nMovNro = MD.nMovNro " & _
'       "LEFT JOIN (SELECT nMovNro, nMovItem, SUM(nMovOtroImporte) as nMovOtroItemImporte " & _
'       "           FROM MOVOTROSITEM WHERE cMovOtroVariable = '21140102' " & _
'       "           GROUP BY nMovNro, nMovItem) AS MOTI ON MC.nMovNro = MOTI.nMovNro and MC.nMovItem = MOTI.nMovItem " & _
'       "LEFT JOIN (SELECT nMovNro, nMovItem, SUM(nMovOtroImporte) as nMovOtroItemImporte " & _
'       "           FROM MOVOTROSITEM WHERE not cMovOtroVariable = '21140102' " & _
'       "           GROUP BY nMovNro, nMovItem) AS MOTO ON MC.nMovNro = MOTO.nMovNro and MC.nMovItem = MOTO.nMovItem " & _
'       "     JOIN MovGasto MO ON mo.nMovNro = m.nMovNro " & _
'       "     JOIN OpeDoc OD ON od.nDocTpo = md.nDocTpo " & _
'       "LEFT JOIN MovTpoCambio mtc ON mtc.nMovNro = m.nMovNro " & _
'       "LEFT JOIN Persona P ON MO.cPersCod = P.cPersCod " & _
'       "LEFT JOIN PersID pid ON pid.cPersCod = P.cPersCod and pid.cPersIDTpo = '2' " & _
'       "WHERE M.nMovEstado = '10' and M.nMovFlag not IN ('1','3','2','5') " & _
'       " and CONVERT(varchar(8),md.dDocFecha,112) LIKE '200705%' and OD.cOpeCod = '700101' " & _
'        "and MC.nMovImporte > 0 GROUP BY P.cPerscod, P.cPersNombre, pid.cPersIDnro " & _
'        "HAVING SUM(MC.nMovImporte) >= 330.00 ORDER BY P.cPersNombre "
sSql = "stp_sel_CargaCOAProveedores '" & dFecha & "'"
Set CargaCOAProveedores = dbConec.CargaRecordSet(sSql)
End Function

Public Function CargaCOADocumentos(psPeriodo As String, psCtaIgv As String, psOpeCod As String, Optional psDocTpo As String = "") As Recordset
'**ALPA 26032008**************************************************************************************************************************************
'sSQL = "SELECT M.cMovNro,M.nMovNro,M.cOpeCod, MD.nDocTpo,ISNULL(P.cPersNombre,'') cPersNombre, ISNULL(pid.cPersIDnro,'') as cProvRuc,MD.dDocFecha,  MD.cDocNro, ISNULL(SUM(MOTI.nMovOtroItemImporte),0) nIGV, ISNULL(SUM(MOTO.nMovOtroItemImporte),0)  nOtros, sum(MC.nMovImporte) as nMovImporte, SUM(me.nMovMEImporte) as nMovMEImporte, ISNULL(mtc.nMovTpoCambio,1) nTC  " _
'     & "FROM MOV M JOIN MOVCTA MC ON M.nMovNro = MC.nMovNro LEFT JOIN MovME me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'     & "           JOIN MOVDOC MD ON M.nMovNro = MD.nMovNro " _
'     & "      LEFT JOIN (SELECT nMovNro, nMovItem, SUM(nMovOtroImporte) as nMovOtroItemImporte " _
'     & "                 FROM MOVOTROSITEM WHERE cMovOtroVariable = '" & psCtaIgv & "' " _
'     & "                 GROUP BY nMovNro, nMovItem " _
'     & "                ) AS MOTI ON MC.nMovNro = MOTI.nMovNro and MC.nMovItem = MOTI.nMovItem " _
'     & "      LEFT JOIN (SELECT nMovNro, nMovItem, SUM(nMovOtroImporte) as nMovOtroItemImporte " _
'     & "                 FROM MOVOTROSITEM WHERE not cMovOtroVariable = '" & psCtaIgv & "' " _
'     & "                 GROUP BY nMovNro, nMovItem " _
'     & "                ) AS MOTO ON MC.nMovNro = MOTO.nMovNro and MC.nMovItem = MOTO.nMovItem " _
'     & "           JOIN MovGasto MO ON mo.nMovNro = m.nMovNro " _
'     & "           JOIN OpeDoc OD ON od.nDocTpo = md.nDocTpo " _
'     & "      LEFT JOIN MovTpoCambio mtc ON mtc.nMovNro = m.nMovNro " _
'     & "      LEFT JOIN Persona P ON MO.cPersCod = P.cPersCod " _
'     & "      LEFT JOIN PersID pid ON pid.cPersCod = P.cPersCod and pid.cPersIDTpo = '" & gPersIdRUC & "' " _
'     & "WHERE M.nMovEstado = '" & gMovEstContabMovContable & "' and M.nMovFlag not IN ('" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagExtornado & "','" & gMovFlagModificado & "') " _
'     & "  and Substring(M.cMovNro,1,6) LIKE '" & psPeriodo & "%' and OD.cOpeCod = '" & psOpeCod & "' and MC.nMovImporte > 0 " _
'     & IIf(psDocTpo = "", "", " and md.nDocTpo IN (" & psDocTpo & ") ") _
'     & "GROUP BY M.cMovNro,M.nMovNro,M.cOpeCod,MD.nDocTpo,P.cPersNombre, pid.cPersIDnro,MD.dDocFecha, MD.cDocNro, ISNULL(mtc.nMovTpoCambio,1) HAVING SUM(MC.nMovImporte) >= 330.00 ORDER BY P.cPersNombre, md.cDocNro "
'**End***************************************************************************************************************************************************
'     'CONVERT(varchar(8),md.dDocFecha,112) LIKE
sSql = "stp_sel_ReportesTransferenciaComprobantes '" & psPeriodo & "','" & psCtaIgv & "','" & gPersIdRUC & "','" & gMovEstContabMovContable & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagExtornado & "','" & gMovFlagModificado & "', '" & psOpeCod & "', '" & psDocTpo & "'"
Set CargaCOADocumentos = dbConec.CargaRecordSet(sSql)

End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oIni = New ClasIni
vsBaseComunes = oIni.BaseComunes
vsBasePersonas = oIni.BasePersonas
sLPT = "LPT1"
BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")

Set oIni = Nothing
Set dbConec = New DConecta
dbConec.AbreConexion
End Sub

Private Sub Class_Terminate()
dbConec.CierraConexion
Set dbConec = Nothing
End Sub

'Se quito la condicion de Cuentas contables by gitu 14-01-2009
Public Function GetCompraGastosConsolidado(psDocs As String, pdFechaDel As Date, pdFechaAl As Date, sCtaIGV As String, sCtaISC As String, Optional plFiltroDoc As Boolean = False, Optional pbOrdDoc As Boolean = False, _
                                        Optional sCtaRetrac As String = "", Optional sCtaRetCuarta As String) As Recordset

Dim oImpuesto As DImpuesto
Set oImpuesto = New DImpuesto
Dim lsCtaCVPerdida As String
Dim oCon As NConstSistemas
Set oCon = New NConstSistemas

Dim lnIGVValor As Currency
lnIGVValor = 1 + (oImpuesto.CargaImpuestoFechaValor(oCon.LeeConstSistema(40), CDate(oCon.LeeConstSistema(15))) / 100)
lsCtaCVPerdida = oCon.LeeConstSistema(21)

Dim sObj As String, sImp As String, sDoc As String
    If Not pbOrdDoc Then
        sObj = " LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '" & sCtaIGV & "') as OtroItmIGV ON (OtroItmIGV.nMovNro = a.nMovNro and OtroItmIGV.nMovItem = a.nMovItem) " _
             & " LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '" & sCtaISC & "') as OtroItmISC ON (OtroItmISC.nMovNro = a.nMovNro and OtroItmISC.nMovItem = a.nMovItem) " _
             & " LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '" & sCtaRetrac & "') as OtroItmRetrac ON (OtroItmRetrac.nMovNro = a.nMovNro and OtroItmRetrac.nMovItem = a.nMovItem)  "
             
        If Len(psDocs) > 0 Then
           sDoc = IIf(plFiltroDoc, "", " LEFT") & " JOIN (SELECT * FROM MovDoc   WHERE nDocTpo IN (" & psDocs & ") ) as Doc ON (Doc.nMovNro = m.nMovNro) "
        End If
        sSql = "SELECT c.cCtaContCod, c.cCtaContDesc, ISNULL(SUM(z.nTotPV),0) as nPV, ISNULL(SUM(CASE WHEN Substring(z.cOpeCod,3,1) <> '2' THEN  z.nTotIGV WHEN SubString(z.cOpeCod,3,1) = '2' and z.nTotIGV <> 0 THEN Round(Z.nTotPV - (z.nTotPV / " & lnIGVValor & "), 2) END ),0) as nIGV, ISNULL(SUM(z.nTotISC),0) as nISC " _
             & "FROM   " & vsBaseComunes & "CtaCont c, " _
             & "     ( SELECT a.cCtaContCod, m.cOpeCod, ISNULL(SUM(a.nMovImporte),0) as nTotPV, SUM(ISNULL(OtroItmIGV.nMovOtroImporte,0)) as nTotIGV, SUM(ISNULL(OtroItmISC.nMovOtroImporte,0)) as nTotISC" _
             & "       FROM   Mov m JOIN MovCta a ON m.nMovNro = a.nMovNro " & sObj & sImp & sDoc & " " _
             & "       WHERE  m.nMovEstado = '" & gMovEstContabMovContable & "' and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & "," & gMovFlagExtornado & ") and " _
             & "             substring(m.cMovNro,1,8) BETWEEN '" & Format(pdFechaDel, "yyyymmdd") & "' and '" & Format(pdFechaAl, "yyyymmdd") & "' " _
             & "       GROUP BY a.cCtaContCod, m.cOpeCod " _
             & "     ) as z " _
             & "WHERE  z.cCtaContCod LIKE c.cCtaContCod + '%' " _
             & "GROUP BY c.cCtaContCod, c.cCtaContDesc ORDER BY c.cCtaContCod"
    Else
        sObj = " " _
             & " LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '" & sCtaIGV & "') as OtroItmIGV ON (OtroItmIGV.nMovNro = a.nMovNro and OtroItmIGV.nMovItem = a.nMovItem) " _
             & " LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '" & sCtaISC & "') as OtroItmISC ON (OtroItmISC.nMovNro = a.nMovNro and OtroItmISC.nMovItem = a.nMovItem) " _
             & " LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '" & sCtaRetrac & "') as OtroItmRetrac ON (OtroItmRetrac.nMovNro = a.nMovNro and OtroItmRetrac.nMovItem = a.nMovItem)  " _
             & " LEFT JOIN (SELECT * FROM MovOtrosItem WHERE cMovOtroVariable = '" & sCtaRetCuarta & "') as OtroItmRentaCuarta ON (OtroItmRentaCuarta.nMovNro = a.nMovNro and OtroItmRentaCuarta.nMovItem = a.nMovItem)  "
             
        If Len(psDocs) > 0 Then
           sDoc = IIf(plFiltroDoc, "", " INNER") & " JOIN (SELECT * FROM MovDoc   WHERE nDocTpo IN (" & psDocs & ") ) as Doc ON (Doc.nMovNro = m.nMovNro) "
        End If
        'Se quito la condicion de cuentas contables and a.cCtaContCod IN (" & psCtaS & ")
        'By Gitu 14-01-2009
        sSql = "Select  w.nDocTpo Tipo,w.cDocDesc Descripcion ,SUM(w.nPV) Total,SUM(w.Renta4)Renta,Count(*) cantidad FROM (" _
             & "SELECT  z.nDocTpo, z.cDocDesc,Tipo=z.cDocAbrev,ISNULL(SUM(z.nTotPV),0) as nPV,isnull(sum(z.nTotRentaCuarta),0)as Renta4 " _
             & " From " _
             & "     ( SELECT Doc.cDocNro, m.nMovNro, MTC.nMovTpoCambio , m.cOpeCod , Doc.nDocTpo, Doc.dDocFecha  , m.cMovNro, m.cMovDesc,Docum.cDocAbrev,Docum.cDocDesc, ISNULL(SUM(a.nMovImporte),0) as nTotPV, SUM(ISNULL(OtroItmIGV.nMovOtroImporte,0)) as nTotIGV, SUM(ISNULL(OtroItmISC.nMovOtroImporte,0)) as nTotISC, SUM(ISNULL(OtroItmRetrac.nMovOtroImporte,0)) as nTotRetrac,SUM(ISNULL(OtroItmRentaCuarta.nMovOtroImporte,0)) as nTotRentaCuarta,ISNULL(SUM(ME.nMovMEImporte),0) as nTotPVME  " _
             & "       FROM   Mov m JOIN MovCta a ON m.nMovNro = a.nMovNro " & sObj & sImp & sDoc & " INNER JOIN Documento Docum ON Docum.nDocTpo=Doc.nDocTpo " _
             & "       Left Join MovME ME On m.nMovNro = ME.nMovNro and a.nMovItem = me.nMovItem " _
             & "       Left Join MovTpoCambio MTC On m.nMovNro = MTC.nMovNro " _
             & "       WHERE  m.nMovEstado = '" & gMovEstContabMovContable & "' and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") AND A.nMovImporte > 0 and " _
             & "             substring(m.cMovNro,1,8) BETWEEN '" & Format(pdFechaDel, "yyyymmdd") & "' and '" & Format(pdFechaAl, "yyyymmdd") & "' and a.cctacontcod NOT Like '" & Trim(lsCtaCVPerdida) & "%' " _
             & "       GROUP BY Doc.cDocNro, m.cMovNro,MTC.nMovTpoCambio, m.cMovDesc, m.cOpeCod, m.nMovNro, Doc.nDocTpo, Doc.dDocFecha, Docum.cDocAbrev ,Docum.cDocDesc " _
             & "     ) as z " _
             & " GROUP BY z.cDocNro, z.cMovNro, z.cMovDesc, z.cOpeCod, z.nMovNro, z.nDocTpo,z.cDocAbrev,z.cDocDesc " _
             & " ) as w " _
             & " GROUP BY   w.nDocTpo,w.cDocDesc " _
             & " Order By w.nDocTpo"
    End If
    
    '*** PEAC 20110831
    sSql = "exec stp_sel_GetCompraGastosConsolidado '" & Format(pdFechaDel, "yyyymmdd") & "','" & Format(pdFechaAl, "yyyymmdd") & "','" & Trim(lsCtaCVPerdida) & "'"
    
    Set GetCompraGastosConsolidado = dbConec.CargaRecordSet(sSql)
End Function

Public Function GetMovCompraGastosMayorRef(ByVal pnMovNro As Long) As Recordset
Dim sSql As String

   sSql = " Select isnull(mr.nMovnroRef,'')nMovnroRef , isnull(mdo.cNroConstancia,'')cDocNro,isnull(mdo.dFecha,'')dDocFecha " & _
          " from Mov m " & _
          " inner join MovRef mr ON mr.nmovnro =m.nmovnro " & _
          " inner join MovDocOtros mdo ON mdo.nmovnro =m.nmovnro" & _
          " inner join MovDoc md ON md.nmovnro =m.nmovnro" & _
          " inner join Documento d ON d.nDocTpo = md.nDocTpo and md.nDocTpo = 47 " & _
          " where mr.nmovnroref = " & pnMovNro & " and  m.nmovflag not in (1,2,3) "
    
    Set GetMovCompraGastosMayorRef = dbConec.CargaRecordSet(sSql)
End Function

Public Function GetRegistroComprasExcel() As Recordset
Dim sSql As String

   sSql = "Select * from RegComExcel"
    
    Set GetRegistroComprasExcel = dbConec.CargaRecordSet(sSql)
End Function

Public Function GetDaotIngresos(ByVal psAnio As String, ByVal pnTipCamb As Double) As Recordset
Dim sSql As String

   sSql = " Select M.cPerscod, " _
        & " Case P.nPersPersoneria When '1' Then '01' Else '02' End As Personeria, " _
        & " Case PD.cPersIDtpo When 2 Then '6' When 3 Then '2' When 11 Then '7' When 99 Then '-' Else Cast(PD.cPersIDtpo as varchar(10)) End As TpoDoc, " _
        & " Sum(M.Monto) Monto, PD.cPersIdNro, replace(P.cpersNombre,'-\','') as cPersNombre " _
        & " From(Select PP.cPerscod, MC.cCtaCod, " _
        & "          Case When Substring(MC.cCtaCod,9,1) = 2 then (Sum(MC.nMonto)* " & pnTipCamb & ") else Sum(MC.nMonto) End As Monto " _
        & "      from MovColDet MC " _
        & "          join Mov M On MC.nMovNro = M.nMovNro " _
        & "          join ProductoPersona PP On MC.cCtaCod = PP.cCtaCod and PP.nPrdPersRelac = 20 " _
        & "      Where MC.nPrdConceptoCod Like '[123]10%' and Left(M.cMovnro,4) = '" & psAnio & "'" _
        & "         And M.nMovFlag not in (" & gMovFlagEliminado & "," & gMovFlagExtornado & ")" _
        & "      Group By PP.cPerscod, MC.cCtaCod) As M " _
        & "     Join Persona P On M.cPerscod = P.cPerscod " _
        & "     Join PersId PD On P.cPerscod = PD.cPerscod And PD.cPersIdTpo = (Select Min(cPersIDtpo) from Persid A Where A.cPerscod = PD.cPerscod) " _
        & " Group By M.cPerscod, P.cpersNombre, PD.cPersidNro, PD.cPersIDtpo, Case P.nPersPersoneria When '1' Then '01' Else '02' End " _
        & " Having Sum(m.Monto) > 6900.00 order by Monto"
    
    Set GetDaotIngresos = dbConec.CargaRecordSet(sSql)
End Function

Public Function GetDaotCostos(ByVal psAnio As String, ByVal pnTipCamb As Double) As Recordset
Dim sSql As String

   sSql = " Select M.cPerscod, Case P.nPersPersoneria When '1' Then '01' Else '02' End As Personeria, " _
        & " Case PD.cPersIDtpo When 2 Then '6' When 3 Then '2' When 11 Then '7' When 99 Then '-' Else Cast(PD.cPersIDtpo as varchar(10)) End As TpoDoc, " _
        & " PD.cPersIdNro, Sum(M.Monto) Monto, replace(P.cpersNombre,'-\','') as cPersNombre" _
        & " From(Select PP.cPerscod, MC.cCtaCod, " _
        & "         Case When Substring(MC.cCtaCod,9,1) = 2 then (Sum(MC.nMonto)* 2.996) else Sum(MC.nMonto) End As Monto " _
        & "      from MovCapDet MC " _
        & "         join Mov M On MC.nMovNro = M.nMovNro " _
        & "         join ProductoPersona PP On MC.cCtaCod = PP.cCtaCod and PP.nPrdPersRelac = 10 " _
        & "      Where Left(M.cMovNro,4) = '" & psAnio & "'" _
        & "          And MC.cOpecod In ('200220','200500','200501','200502','200503','210200','210201','210202','210203','210206','210401','210806','220701','401626','402526','402626') " _
        & "          And M.nMovFlag not in (" & gMovFlagEliminado & "," & gMovFlagExtornado & ")" _
        & "      Group By PP.cPerscod, MC.cCtaCod) As M " _
        & "     Join Persona P On M.cPerscod = P.cPerscod " _
        & "     Join PersId PD On P.cPerscod = PD.cPerscod And PD.cPersIdTpo = (Select Min(cPersIDtpo) from Persid A Where A.cPerscod = PD.cPerscod) " _
        & " Where cPersIDtpo <> 10 " _
        & " group By M.cPerscod, P.cpersNombre, PD.cPersidNro, PD.cPersIDtpo, Case P.nPersPersoneria When '1' Then '01' Else '02' End " _
        & " Having Sum(m.Monto) > 6900.00 " _
        & " Union All " _
        & " Select MG.cPerscod, Case P.nPersPersoneria When '1' Then '01' Else '02' End As Personeria, " _
        & "     Case PD.cPersIDtpo When '2' Then '6' When '3' Then '2' When '11' Then '7' Else Cast(PD.cPersIDtpo as varchar(10)) End As TpoDoc, " _
        & "     PD.cPersIdNro, Sum(MC.nMovImporte) Monto, P.cpersNombre " _
        & " from MovCta MC " _
        & "     Join Mov M On MC.nMovNro = M.nMovNro " _
        & "     Join MovDoc MD On MC.nMovNro = MD.nMovNro " _
        & "     Join MovGasto MG On MC.nMovNro = MG.nMovNro "
    sSql = sSql & " Join Persona P On MG.cPerscod = P.cPerscod " _
        & "     join Persid PD On P.cPerscod = PD.cPerscod And PD.cPersIdTpo = (Select Min(cPersIDtpo) from Persid A Where A.cPerscod = PD.cPerscod) " _
        & " Where Left(M.cMovNro,4) = '" & psAnio & "'" _
        & "     and MD.nDocTpo In (2,3) And MC.nMovImporte > 0 " _
        & "     And M.nMovFlag not in (" & gMovFlagEliminado & "," & gMovFlagExtornado & ")" _
        & " group by MG.cPerscod, P.cpersNombre, PD.cPersidNro, PD.cPersIDtpo, Case P.nPersPersoneria When '1' Then '01' Else '02' End " _
        & " Having Sum(MC.nMovImporte) > 6900.00 order by Monto"
    
    Set GetDaotCostos = dbConec.CargaRecordSet(sSql)
End Function

Public Function GetGastosAdministrativosDatos(ByVal psAgeCod As String, ByVal psAreaCod As String, _
                                       ByVal psFechaIni As String, ByVal psFechaFin As String, _
                                       ByVal psCtaS As String, ByVal psDocs As String) As ADODB.Recordset
'*******************************************
'* Create by GITU Según Acta 195-2008/TI-D *
'*******************************************
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsFiltroAge  As String
Dim lsFiltroArea  As String
Dim lsJoinArea As String
Dim lsFecAnt As String

Set oConect = New DConecta
Set rs = New Recordset

lsFecAnt = Format(CDate(psFechaIni) - 1, gsFormatoFecha)
psFechaFin = Mid(psFechaFin, 7, 4) + Mid(psFechaFin, 4, 2) + Mid(psFechaFin, 1, 2)
psFechaIni = Mid(psFechaIni, 7, 4) + Mid(psFechaIni, 4, 2) + Mid(psFechaIni, 1, 2)


If oConect.AbreConexion = False Then Exit Function

If psAgeCod = "" Then
    lsFiltroAge = ""
Else
    lsFiltroAge = " and Right(a.cCtaContCod,2) In ('" & psAgeCod & "') "
End If
If psAreaCod = "" Or Len(psAreaCod) > 3 Then
    lsFiltroArea = ""
Else
    lsFiltroArea = " and MA.cAreaCod In ('" & psAreaCod & "') "
    lsJoinArea = "LEFT JOIN MovObjAreaAgencia MA On b.nMovNro = MA.nMovNro "
End If

sql = "SELECT Left(a.cCtaContCod,2) + '0' + Substring(a.cCtaContCod,4,25) as cCtaContCod, " _
    & "       Doc.dDocFecha, Doc.cDocNro, Obj.cPersNombre, b.cMovDesc, " _
    & "       dbo.getsaldocta('" & lsFecAnt & "',a.cCtaContCod,1) As nSalAnt, " _
    & "       ISNULL(Sum(a.nMovImporte), 0) As NPV " _
    & "FROM Mov b JOIN MovCta a ON a.nMovNro = b.nMovNro " _
    & "    LEFT JOIN (SELECT mg.nMovNro, per.cPersNombre, ISNULL(pid.cPersIDnro,'') cRuc " _
    & "           FROM MovGasto mg " _
    & "                JOIN Persona Per ON Per.cPersCod = mg.cPersCod " _
    & "                LEFT JOIN ( SELECT * FROM PersID WHERE cPersIDTpo = '2' ) pid ON pid.cPersCod = per.cPersCod) as Obj ON Obj.nMovNro = a.nMovNro " _
    & "    LEFT JOIN (SELECT * FROM MovDoc WHERE nDocTpo IN (" & psDocs & ") ) as Doc ON (Doc.nMovNro = b.nMovNro) " _
    & "    LEFT JOIN MovRef mr ON mr.nMovNro = b.nMovNro " _
    & lsJoinArea _
    & "    LEFT JOIN MOVDOC mdChq ON mdChq.nMovNro = mr.nMovNroRef and mdChq.nDocTpo IN (47 ) , CtaCont c " _
    & "WHERE  b.nMovEstado = '10' and not b.nMovFlag in (1,5,2) " _
    & lsFiltroAge & lsFiltroArea _
    & "    and a.cCtaContCod IN (" & psCtaS & ") " _
    & "    and c.cCtaContCod = a.cCtaContCod and substring(b.cMovNro,1,8) BETWEEN '" & psFechaIni & "' and '" & psFechaFin & " ' " _
    & "GROUP BY a.cCtaContCod, Doc.dDocFecha, Doc.cDocNro, Obj.cPersNombre, b.cMovDesc " _
    & "ORDER BY Left(a.cCtaContCod,2) + '0' + Substring(a.cCtaContCod,4,25)"
    
Set rs = oConect.CargaRecordSet(sql)
Set GetGastosAdministrativosDatos = rs
oConect.CierraConexion
Set oConect = Nothing
End Function
'EJVG20130402 ***
Public Function ExisteRegistroVentasPLE(ByVal pdFecha As Date) As Boolean
    Dim oConect As New DConecta
    Dim rs As New ADODB.Recordset
    Dim lsSql As String

    On Error GoTo ErrExisteRegistroCompras
    lsSql = "SELECT TOP 1 * " _
        & "  FROM PLE_RegistroVentaDet D INNER JOIN PLE_RegistroVenta C ON C.nId = D.nId " _
        & "  WHERE DATEDIFF(D,dFecha,'" & Format(pdFecha, "yyyymmdd") & "')=0 " _
        & "     AND dFechaGenera = (SELECT MAX(dFechaGenera) FROM PLE_RegistroVenta C1 WHERE DATEDIFF(D,C1.dFecha,C.dFecha)=0) "
    
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(lsSql)
    If rs.RecordCount > 0 Then
        ExisteRegistroVentasPLE = True
    Else
        ExisteRegistroVentasPLE = False
    End If
    oConect.CierraConexion
    Set rs = Nothing
    Set oConect = Nothing
    Exit Function
ErrExisteRegistroCompras:
    Err.Raise Err.Number, "NContImpreReg:ExisteRegistroCompras", Err.Description
End Function
Public Function ExisteRegistroComprasPLE(ByVal pdFecha As Date) As Boolean
    Dim oConect As New DConecta
    Dim rs As New ADODB.Recordset
    Dim lsSql As String

    On Error GoTo ErrExisteRegistroCompras
    lsSql = "SELECT TOP 1 * " _
        & "  FROM PLE_RegistroCompraDet D INNER JOIN PLE_RegistroCompra C ON C.nId = D.nId " _
        & "  WHERE DATEDIFF(D,dFecha,'" & Format(pdFecha, "yyyymmdd") & "')=0 " _
        & "     AND dFechaGenera = (SELECT MAX(dFechaGenera) FROM PLE_RegistroCompra C1 WHERE DATEDIFF(D,C1.dFecha,C.dFecha)=0) "
    
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(lsSql)
    If rs.RecordCount > 0 Then
        ExisteRegistroComprasPLE = True
    Else
        ExisteRegistroComprasPLE = False
    End If
    oConect.CierraConexion
    Set rs = Nothing
    Set oConect = Nothing
    Exit Function
ErrExisteRegistroCompras:
    Err.Raise Err.Number, "NContImpreReg:ExisteRegistroCompras", Err.Description
End Function
Public Sub InsertaRegistroVentasPLE(ByVal pdFecha As Date, ByVal pdFechaGenera As Date, psUsuarioGenera As String)
    Dim oConect As New DConecta
    Dim lsSql As String

    On Error GoTo ErrInsertaRegistroVentasPLE
    lsSql = "Exec stp_ins_PLE_RegistroVenta '" & Format(pdFecha, "yyyymmdd") & "','" & Format(pdFechaGenera, "yyyymmdd hh:mm:ss") & "','" & psUsuarioGenera & "'"
    oConect.AbreConexion
    oConect.Ejecutar (lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Sub
ErrInsertaRegistroVentasPLE:
    Err.Raise Err.Number, "NContImpreReg:InsertaRegistroVentasPLE", Err.Description
End Sub
Public Sub InsertaRegistroComprasPLE(ByVal pdFecha As Date, ByVal pdFechaGenera As Date, psUsuarioGenera As String)
    Dim oConect As New DConecta
    Dim lsSql As String

    On Error GoTo ErrInsertaRegistroComprasPLE
    lsSql = "Exec stp_ins_PLE_RegistroCompra '" & Format(pdFecha, "yyyymmdd") & "','" & Format(pdFechaGenera, "yyyymmdd hh:mm:ss") & "','" & psUsuarioGenera & "'"
    oConect.AbreConexion
    oConect.Ejecutar (lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Sub
ErrInsertaRegistroComprasPLE:
    Err.Raise Err.Number, "NContImpreReg:InsertaRegistroComprasPLE", Err.Description
End Sub
Public Function RecuperaDatosRegistroComprasPLE(ByVal pdFecha As Date) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim lsSql As String

    On Error GoTo ErrRecuperaDatosRegistroComprasPLE
    lsSql = "Exec stp_sel_PLE_RegistroCompra '" & Format(pdFecha, "yyyymmdd") & "'"
    oConect.AbreConexion
    Set RecuperaDatosRegistroComprasPLE = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
ErrRecuperaDatosRegistroComprasPLE:
    Err.Raise Err.Number, "NContImpreReg:RecuperaDatosRegistroCompras", Err.Description
End Function
Public Function RecuperaDatosRegistroVentasPLE(ByVal pdFecha As Date) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim lsSql As String
    
    On Error GoTo ErrRecuperaDatosRegistroVentasPLE
    lsSql = "Exec stp_sel_PLE_RegistroVenta '" & Format(pdFecha, "yyyymmdd") & "'"
    oConect.AbreConexion
    Set RecuperaDatosRegistroVentasPLE = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
ErrRecuperaDatosRegistroVentasPLE:
    Err.Raise Err.Number, "NContImpreReg:RecuperaDatosRegistroVentas", Err.Description
End Function
'END EJVG *******
'ALPA 20130529***
Public Function ExisteLibroDiarioPLE(ByVal pdFecha As Date) As Boolean
    Dim oConect As New DConecta
    Dim rs As New ADODB.Recordset
    Dim lsSql As String

    On Error GoTo ErrExisteLibroDiarioPLE
    lsSql = "SELECT TOP 1 * " _
        & "  FROM PLE_LibroDiarioDet D INNER JOIN PLE_LibroDiario C ON C.nId = D.nId " _
        & "  WHERE DATEDIFF(D,dFecha,'" & Format(pdFecha, "yyyymmdd") & "')=0 " _
        & "     AND dFechaGenera = (SELECT MAX(dFechaGenera) FROM PLE_LibroDiario C1 WHERE DATEDIFF(D,C1.cPeriodo,C.cPeriodo)=0) "
    
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(lsSql)
    If rs.RecordCount > 0 Then
        ExisteLibroDiarioPLE = True
    Else
        ExisteLibroDiarioPLE = False
    End If
    oConect.CierraConexion
    Set rs = Nothing
    Set oConect = Nothing
    Exit Function
ErrExisteLibroDiarioPLE:
    Err.Raise Err.Number, "NContImpreReg:ExisteLibroDiarioPLE", Err.Description
End Function

Public Sub InsertaLibroDiarioPLE(ByVal pdFecha As Date, ByVal pdFechaGenera As Date, psUsuarioGenera As String)
    Dim oConect As New DConecta
    Dim lsSql As String

    On Error GoTo ErrInsertaLibroDiarioPLE
    lsSql = "Exec PLE_stp_ins_LibroDiario '" & CStr(Format(pdFecha, "yyyymmdd")) & "','" & Format(pdFechaGenera, "yyyymmdd hh:mm:ss") & "','" & psUsuarioGenera & "'"
    oConect.AbreConexion
    oConect.Ejecutar (lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Sub
ErrInsertaLibroDiarioPLE:
    Err.Raise Err.Number, "NContImpreReg:InsertaLibroDiarioPLE", Err.Description
End Sub
Public Function RecuperaDatosLibroDiarioPLE(ByVal pdFecha As Date) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim lsSql As String
    
    On Error GoTo ErrRecuperaDatosRegistroVentasPLE
    lsSql = "Exec PLE_stp_sel_LibroDiario '" & CStr(Format(pdFecha, "yyyymmdd")) & "'"
    oConect.AbreConexion
    Set RecuperaDatosLibroDiarioPLE = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
ErrRecuperaDatosRegistroVentasPLE:
    Err.Raise Err.Number, "NContImpreReg:RecuperaDatosLibroDiarioPLE", Err.Description
End Function
'PASI 20140531
Public Function RecuperaDatosLibroDiarioDetCtaContPLE(ByVal pdFecha As Date) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim lsSql As String
    
    On Error GoTo ErrRecuperaDatosLibroDiarioDet
    lsSql = "Exec Ple_stp_sel_LibroDiarioDetCtaCont '" & CStr(Format(pdFecha, "yyyymmdd")) & "'"
    oConect.AbreConexion
    Set RecuperaDatosLibroDiarioDetCtaContPLE = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
ErrRecuperaDatosLibroDiarioDet:
    Err.Raise Err.Number, "NContImpreReg:ErrRecuperaDatosLibroDiarioDet", Err.Description
End Function
'end PASI
Public Function ExisteLibroMayorPLE(ByVal pdFecha As Date) As Boolean
    Dim oConect As New DConecta
    Dim rs As New ADODB.Recordset
    Dim lsSql As String

    On Error GoTo ErrExisteLibroMayorPLE
    lsSql = "SELECT TOP 1 * " _
        & "  FROM PLE_LibroMayorDet D INNER JOIN PLE_LibroMayor C ON C.nId = D.nId " _
        & "  WHERE DATEDIFF(D,dFecha,'" & Format(pdFecha, "yyyymmdd") & "')=0 " _
        & "     AND dFechaGenera = (SELECT MAX(dFechaGenera) FROM PLE_LibroMayor C1 WHERE DATEDIFF(D,C1.cPeriodo,C.cPeriodo)=0) "
    
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(lsSql)
    If rs.RecordCount > 0 Then
        ExisteLibroMayorPLE = True
    Else
        ExisteLibroMayorPLE = False
    End If
    oConect.CierraConexion
    Set rs = Nothing
    Set oConect = Nothing
    Exit Function
ErrExisteLibroMayorPLE:
    Err.Raise Err.Number, "NContImpreReg:ExisteLibroMayorPLE", Err.Description
End Function

Public Sub InsertaLibroMayorPLE(ByVal pdFecha As Date, ByVal pdFechaGenera As Date, psUsuarioGenera As String)
    Dim oConect As New DConecta
    Dim lsSql As String

    On Error GoTo ErrInsertaLibroMayorPLE
    lsSql = "Exec PLE_stp_ins_LibroMayor '" & CStr(Format(pdFecha, "yyyymmdd")) & "','" & Format(pdFechaGenera, "yyyymmdd hh:mm:ss") & "','" & psUsuarioGenera & "'"
    oConect.AbreConexion
    oConect.Ejecutar (lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Sub
ErrInsertaLibroMayorPLE:
    Err.Raise Err.Number, "NContImpreReg:InsertaLibroMayorPLE", Err.Description
End Sub
Public Function RecuperaDatosLibroMayorPLE(ByVal pdFecha As Date) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim lsSql As String
    
    On Error GoTo ErrRecuperaDatosLibroMayorPLE
    lsSql = "Exec PLE_stp_sel_LibroMayor '" & CStr(Format(pdFecha, "yyyymmdd")) & "'"
    oConect.AbreConexion
    Set RecuperaDatosLibroMayorPLE = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
ErrRecuperaDatosLibroMayorPLE:
    Err.Raise Err.Number, "NContImpreReg:RecuperaDatosLibroMayorPLE", Err.Description
End Function
'****************

    
