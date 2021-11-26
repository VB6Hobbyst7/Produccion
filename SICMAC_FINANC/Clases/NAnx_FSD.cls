VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NAnx_FSD"
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
 
Dim sServidorFSD As String
Dim sBaseFSD As String
Dim sServidorConsolidada As String
Dim pnBitCentral As Boolean

Public Event BarraShow(pnMax)
Public Event BarraProgress(value, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
Public Event BarraClose()

Public Function GetFSD_SaldosCont(psCtaCod As String, pdFechaDel As Date, pdFechaAl As Date) As Recordset

   sSql = "SELECT cCtaContCod, dCtaSaldoFecha, nCtaSaldoImporte FROM CtaSaldo " _
        & "WHERE  cCtaContCod LIKE '" & psCtaCod & "%' and " _
        & "       dCtaSaldoFecha BETWEEN '" & Format(pdFechaDel, gsFormatoFecha) & "' and '" & Format(pdFechaAl, gsFormatoFecha) & "' " _
        & "UNION " _
        & "SELECT cCtaContCod, dCtaSaldoFecha, nCtaSaldoImporte FROM CtaSaldo cs " _
        & "WHERE  cCtaContCod LIKE '" & psCtaCod & "%' and dCtaSaldoFecha = " _
        & "      (SELECT Max(dCtaSaldoFecha) FROM CtaSaldo " _
        & "       WHERE  cCtaContCod = cs.cCtaContCod and dCtaSaldoFecha < '" & Format(pdFechaDel, gsFormatoFecha) & "'" _
        & "      ) " _
        & "ORDER BY cCtaContCod, dCtaSaldoFecha"
   Set GetFSD_SaldosCont = dbConec.CargaRecordSet(sSql)
   
End Function

Public Function GetFSD_Estadisticas(pdFechaDel As Date, pnMoneda As Integer, pnActivo As Integer, psProd As String) As Recordset
dbConec.CierraConexion
If dbConec.AbreConexion Then 'Remota("07", True, False, "03")
    sSql = "SELECT * FROM Estad_FSD WHERE FSD_Anio = " & Year(pdFechaDel) & " and FSD_Mes = " & Month(pdFechaDel) & " and SubString(cCodCta,3,4) = '" & psProd & pnMoneda & "' and nActiva = " & pnActivo _
         & " ORDER BY cCodCta "
    Set GetFSD_Estadisticas = dbConec.CargaRecordSet(sSql)
End If
dbConec.CierraConexion
dbConec.AbreConexion

End Function

Public Function GetFSD_EstadisticasNew(pdFechaDel As Date, pnMoneda As Integer, pnActivo As Integer, psProd As String, pnTipoInformacion As Integer, pnExonerados As Integer, pnSaldoInteres As Integer) As Recordset
If Not pnBitCentral Then
    gsCodAge = "07"
    dbConec.CierraConexion
    If Not dbConec.AbreConexion Then 'Remota(gsCodAge, , , "03")
        dbConec.AbreConexion
    End If
Else
    
End If
    '1 Sin  Exonerados ni Excluidos
    '2 solo Exonerados
    '3 solo Excluidos
   If pnBitCentral Then
    sSql = "Select CONVERT(varchar(8),CS.dFecha,112) as cFecha, CS.cCtaCod cCodCta, " & _
         IIf(pnSaldoInteres = 1, "SUM(ROUND( CS.nSaldCnt/F.nNumTit,2))", "SUM(ROUND( CS.nInteres/F.nNumTit,2)) ") & " nMonto, " & _
         IIf(pnSaldoInteres = 1, "SUM(F.nSdoDis) nMontoDis, SUM(F.nSdoCob) nMontoCob ", "SUM(F.nIntMes) nMontoDis, SUM(F.nIntCob) nMontoCob") & _
         " FROM CAPSALDOSDIARIOS CS " & _
         " INNER JOIN " & sBaseFSD & "FSDPersonaDet F ON CS.cCtaCod= F.cCtaCod " & _
         " INNER JOIN Persona P on P.cPersCod=F.cPersCod " & _
         " WHERE F.cPersCod " & IIf(pnExonerados = 2 Or pnExonerados = 3, "", "NOT ") & " IN (SELECT E.cPersCod FROM " & sBaseFSD & "FSDPERSONAEXONERADOS E  " & IIf(pnExonerados = 1, "", " WHERE cPersTipExon " & IIf(pnExonerados = 2, "<", ">=") & " 10") & " )" & _
         " AND Substring(CONVERT(varchar(8),CS.dFecha,112),1,4)='" & Mid(pdFechaDel, 7, 4) & "'" & _
         " AND Substring(CONVERT(varchar(8),CS.dFecha,112),5,2)='" & Mid(pdFechaDel, 4, 2) & "'" & _
         IIf(Len(Trim(psProd)) = 0, "", " AND SUBSTRING(CS.cCtaCod, 6,3)='" & psProd & "'") & _
         " AND SUBSTRING(CS.cCtaCod, 9,1)='" & pnMoneda & "'" & _
         " AND F.bInactiva=" & pnActivo & _
         " AND P.nPersPersoneria IN ('1','2') "
    If pnExonerados = 1 Then
        sSql = sSql & " AND F.cPersCod " & IIf(pnTipoInformacion = 1, "NOT ", "") & " IN (SELECT E.cPersCod  FROM " & sBaseFSD & "FSDPersonaClasif e WHERE cPersTipo=2)"
    End If
   Else
   sSql = "Select CONVERT(varchar(8),CS.dFecha,112) as cFecha, CS.cCodCta, " & _
        IIf(pnSaldoInteres = 1, " SUM(ROUND(CS.nSaldCnt/F.nNumTit,2))", " SUM(ROUND(CS.nInteres/F.nNumTit,2)) ") & " nMonto, " & _
        IIf(pnSaldoInteres = 1, "SUM(F.nSdoDis) nMontoDis, SUM(F.nSdoCob) nMontoCob ", "SUM(F.nIntMes) nMontoDis, SUM(F.nIntCob) nMontoCob") & _
        " FROM " & sBaseFSD & "CAPSALDOSDIARIOS CS " & _
        " INNER JOIN " & sBaseFSD & "FSDPersonaDet F ON CS.cCodCta= F.cCodCta " & _
        " WHERE F.cCodPers " & IIf(pnExonerados = 2 Or pnExonerados = 3, "", "NOT ") & " IN (SELECT E.cPersCod FROM " & sBaseFSD & "FSDPERSONAEXONERADOS E  " & IIf(pnExonerados = 1, "", " ") & " )" & _
        " AND Substring(CONVERT(varchar(8),CS.dFecha,112),1,4)='" & Mid(pdFechaDel, 7, 4) & "'" & _
        " AND Substring(CONVERT(varchar(8),CS.dFecha,112),5,2)='" & Mid(pdFechaDel, 4, 2) & "'" & _
        IIf(Len(Trim(psProd)) = 0, "", " AND SUBSTRING(CS.cCodCta, 3,3)='" & psProd & "'") & _
        " AND SUBSTRING(CS.cCodCta, 6,1)='" & pnMoneda & "'" & _
        " AND F.bInactiva=" & pnActivo
         
         
         
    If pnExonerados = 1 Then
        sSql = sSql & " AND F.cCodPers " & IIf(pnTipoInformacion = 1, " NOT ", "") & " IN (SELECT E.cPersCod  FROM  " & sBaseFSD & "FSDPersonaClasif e WHERE cPersTipo=2)"
    End If
    End If
    
    If pnSaldoInteres = 1 Then
        sSql = sSql & " AND CS.nSaldCnt>0"
    ElseIf pnSaldoInteres = 2 Then
        sSql = sSql & " AND CS.nInteres>0"
    End If
   If pnBitCentral Then
    sSql = sSql & " GROUP BY CONVERT(varchar(8),CS.dFecha,112), CS.cCtaCod"
    sSql = sSql & " ORDER BY CS.cCtaCod ASC, CONVERT(varchar(8),CS.dFecha,112) ASC "
   Else
    sSql = sSql & " GROUP BY CONVERT(varchar(8),CS.dFecha,112), CS.cCodCta"
    sSql = sSql & " ORDER BY CS.cCodCta ASC, CONVERT(varchar(8),CS.dFecha,112) ASC "
   End If
    Set GetFSD_EstadisticasNew = dbConec.CargaRecordSet(sSql)
If Not pnBitCentral Then
    dbConec.CierraConexion
    dbConec.AbreConexion
End If
End Function

Public Function GetFSD_Coberturados(psCodCta As String) As Recordset
gsCodAge = "07"
dbConec.CierraConexion

If dbConec.AbreConexion Then 'Remota(gsCodAge, , , "03")
          
    sSql = "select distinct nSdoDis, nIntMes, nSdoCob, nIntCob from " & sBaseFSD & "fsdpersonadet where cCodCta='" & psCodCta & "'"
    
    Set GetFSD_Coberturados = dbConec.CargaRecordSet(sSql)
End If
dbConec.CierraConexion
dbConec.AbreConexion

End Function
 
Public Function FSD_GeneraEstadisticas(pdFechaDel As Date, pnActivo As Integer, pnMonto_FSD As Double, pnTC As Currency, Optional pnBandera As Integer = 1) As Boolean

'pnBandera=1 si es como se estaba ejecutando hasta ahora
'pnBandera=2 si es la nueva version de Pepe

On Error GoTo FSD_GeneraEstadisticasErr
FSD_GeneraEstadisticas = False

If pnBandera = 1 Then  'FSD Actual
'    dbConec.CierraConexion
'    If dbConec.AbreConexionRemota("07", True, False, "03") Then
'        sSQL = "sp_GETFSD_Saldos " & Year(pdFechaDel) & ", " & Month(pdFechaDel) & ", " & pnMonto_FSD & ", " & pnTC
'        dbConec.Ejecutar sSQL
'    End If
'    dbConec.AbreConexion
ElseIf pnBandera = 2 Then 'FSD Nuevo
    
    dbConec.CierraConexion
    If dbConec.AbreConexion Then 'Remota("07", True, False, "03")
        dbConec.CommadTimeOut = 7200
        sSql = " " & sBaseFSD & "sp_GETFSD_Saldos " & Year(pdFechaDel) & ", " & Month(pdFechaDel) & ", " & pnMonto_FSD & ", " & pnTC
        dbConec.Ejecutar sSql
    End If
    dbConec.AbreConexion
End If

FSD_GeneraEstadisticas = True

Exit Function
FSD_GeneraEstadisticasErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetFSD_SaldoCapCta(pdFecha As Date, psCodCta As String, pnTipAct As Currency) As Recordset
sSql = "SELECT TA.cCodCta, TA.dFecTran, ROUND(ISNULL(TA.nSaldCnt*" & pnTipAct & ",0),2) nSaldCnt  " _
    & "FROM TransahoConsol TA " _
    & "WHERE cCodCta = '" & psCodCta & "' AND " _
    & "      DATEDIFF(mm,TA.dFecTran,'" & Format$(pdFecha, "mm/dd/yyyy") & "')= 0 " _
    & "  AND nNumTran IN (SELECT MAX(nNumTran) FROM TransahoConsol TA1 " _
    & "                   WHERE TA1.cCodCta = TA.cCodCta AND datediff(d,TA1.dFectran,TA.dFectran) = 0 " _
    & "                   GROUP BY TA1.cCodCta " _
    & "                  ) " _
    & "And TA.cCodOpe NOT LIKE '2601%' ORDER BY TA.dFecTran"
dbConec.CierraConexion
If dbConec.AbreConexion Then 'Remota("07", True, False, "03")
    Set GetFSD_SaldoCapCta = dbConec.CargaRecordSet(sSql)
End If
dbConec.AbreConexion
End Function


Public Function GetSaldoIniFSD(pdFecha As Date, psCodCta As String, pnTipAct As Currency, Optional pnBitCentral As Boolean = False) As Double


Dim prs As ADODB.Recordset

If pnBitCentral = True Then
    sSql = "SELECT m.cMovNro, mc.cCtaCod, ISNULL(nSaldoContable,0) nSaldCnt " _
         & "FROM MovCap mc JOIN Mov m ON m.nMovNro = mc.nMovNro " _
         & "WHERE mc.cCtaCod = '" & psCodCta & "' and " _
         & "      mc.nMovNro = (SELECT Max(mc1.nMovNro) FROM MovCap mc1 JOIN Mov m1 ON m1.nMovNro = mc1.nMovNro " _
         & "                    WHERE mc1.cCtaCod = mc.cCtaCod and LEFT(m1.cMovNro,8) < '" & Left(Format(pdFecha, "yyyymmdd"), 6) & "')"
    Set prs = dbConec.CargaRecordSet(sSql)

Else
    sSql = "SELECT TA.cCodCta, TA.dFecTran, ROUND(ISNULL(TA.nSaldCnt*" & pnTipAct & ",0),2) nSaldCnt " _
        & "FROM  TransahoConsol TA " _
        & "WHERE cCodCta = '" & psCodCta & "' AND " _
        & "   DATEDIFF(dd,TA.dFecTran,'" & Format$(DateAdd("d", -1, pdFecha), "mm/dd/yyyy") & "')= 0 " _
        & "   AND dFecTran = (SELECT MAX(dFectran) FROM  TransAhoConsol TA1 WHERE " _
        & "                   TA1.cCodCta = '" & psCodCta & "' AND DATEDIFF(dd,TA1.dFecTran,'" & Format$(DateAdd("d", -1, pdFecha), "mm/dd/yyyy") & "')= 0 " _
        & "                   ) " _
        & "ORDER BY TA.dFecTran"
    If dbConec.AbreConexion Then 'Remota("07", , , "03")
        Set prs = dbConec.CargaRecordSet(sSql)
    End If
    dbConec.AbreConexion
End If
If Not prs.EOF Then
   GetSaldoIniFSD = prs!nSaldCnt
Else
   GetSaldoIniFSD = 0
End If
End Function

Public Function GetNroPersonas(pnFSD As Double, pnTC As Currency, Optional pnBandera As Integer = 1) As ADODB.Recordset
'pnBandera=1 si es como se estaba ejecutando hasta ahora
'pnBandera=2 nueva version

If pnBandera = 1 Then
    
    sSql = "SELECT ISNULL(SUM(CASE WHEN nSaldoAhoS >= nSaldoPFS and nSaldoAhoS >= nSaldoAhoD and nSaldoAhoS >= nSaldoPFD THEN 1 END),0) nContAhoS, " _
         & "       ISNULL(SUM(CASE WHEN nSaldoPFS  > nSaldoAhoS and nSaldoPFS  >  nSaldoAhoD and nSaldoPFS  >  nSaldoPFD THEN 1 END),0) nContPFS, " _
         & "       ISNULL(SUM(CASE WHEN nSaldoAhoD >  nSaldoPFS and nSaldoAhoD >  nSaldoAhoS and nSaldoAhoD >  nSaldoPFD THEN 1 END),0) nContAhoD, " _
         & "       ISNULL(SUM(CASE WHEN nSaldoPFD  > nSaldoAhoS and nSaldoPFD  >  nSaldoAhoD and nSaldoPFD  >  nSaldoPFS THEN 1 END),0) nContPFD " _
         & "FROM ( " _
         & "  SELECT    cPersCod, " _
         & "            ISNULL(SUM( CASE WHEN cCtaCod LIKE '__2321%' THEN (nSdoDis+nIntMes)* CASE WHEN SubString(cCtaCod,6,1) = '1' THEN 1 ELSE " & pnTC & " END END),0) nSaldoAhoS, " _
         & "            ISNULL(SUM( CASE WHEN cCtaCod LIKE '__23[34]1%' THEN (nSdoDis+nIntMes)* CASE WHEN SubString(cCtaCod,6,1) = '1' THEN 1 ELSE " & pnTC & " END END),0) nSaldoPFS, " _
         & "            ISNULL(SUM( CASE WHEN cCtaCod LIKE '__2322%' THEN (nSdoDis+nIntMes)* CASE WHEN SubString(cCtaCod,6,1) = '1' THEN 1 ELSE " & pnTC & " END END),0) nSaldoAhoD, " _
         & "            ISNULL(SUM( CASE WHEN cCtaCod LIKE '__23[34]2%' THEN (nSdoDis+nIntMes)* CASE WHEN SubString(cCtaCod,6,1) = '1' THEN 1 ELSE " & pnTC & " END END),0) nSaldoPFD " _
         & "  From      " & sBaseFSD & "fsdPersonaDet " _
         & "  WHERE     cPersCod IN (   SELECT cPersCod " _
         & "                            From " & sBaseFSD & "fsdPersonaDet " _
         & "                            WHERE cPersCod NOT IN (Select cPersCod FROM " & sBaseFSD & "FSDPersonaExonerados) " _
         & "                            GROUP BY cPersCod " _
         & "            HAVING SUM((nSdoDis+nIntMes)* CASE WHEN SubString(cCtaCod,6,1) = '1' THEN 1 ELSE " & pnTC & " END) > " & pnFSD _
         & "        ) " _
         & "  GROUP BY cPersCod " _
         & ") A "

Else
    sSql = "SELECT ISNULL(SUM(CASE WHEN nSaldoAhoS >= nSaldoPFS and nSaldoAhoS >= nSaldoAhoD and nSaldoAhoS >= nSaldoPFD THEN 1 END),0) nContAhoS, " _
         & "       ISNULL(SUM(CASE WHEN nSaldoPFS  > nSaldoAhoS and nSaldoPFS  >  nSaldoAhoD and nSaldoPFS  >  nSaldoPFD THEN 1 END),0) nContPFS, " _
         & "       ISNULL(SUM(CASE WHEN nSaldoAhoD >  nSaldoPFS and nSaldoAhoD >  nSaldoAhoS and nSaldoAhoD >  nSaldoPFD THEN 1 END),0) nContAhoD, " _
         & "       ISNULL(SUM(CASE WHEN nSaldoPFD  > nSaldoAhoS and nSaldoPFD  >  nSaldoAhoD and nSaldoPFD  >  nSaldoPFS THEN 1 END),0) nContPFD " _
         & "FROM ( " _
         & "  SELECT cCodPers, " _
         & "       ISNULL(SUM( CASE WHEN cCodCta LIKE '__2321%' THEN (nSdoDis+nIntMes)* CASE WHEN SubString(cCodCta,6,1) = '1' THEN 1 ELSE " & pnTC & " END END),0) nSaldoAhoS, " _
         & "       ISNULL(SUM( CASE WHEN cCodCta LIKE '__23[34]1%' THEN (nSdoDis+nIntMes)* CASE WHEN SubString(cCodCta,6,1) = '1' THEN 1 ELSE " & pnTC & " END END),0) nSaldoPFS, " _
         & "       ISNULL(SUM( CASE WHEN cCodCta LIKE '__2322%' THEN (nSdoDis+nIntMes)* CASE WHEN SubString(cCodCta,6,1) = '1' THEN 1 ELSE " & pnTC & " END END),0) nSaldoAhoD, " _
         & "       ISNULL(SUM( CASE WHEN cCodCta LIKE '__23[34]2%' THEN (nSdoDis+nIntMes)* CASE WHEN SubString(cCodCta,6,1) = '1' THEN 1 ELSE " & pnTC & " END END),0) nSaldoPFD " _
         & "  From " & sBaseFSD & "fsdPersonaDet " _
         & "  WHERE cCodPers IN (SELECT cCodPers " _
         & "            From " & sBaseFSD & "fsdPersonaDet " _
         & "            WHERE cCodPers NOT IN (Select cPersCod FROM " & sBaseFSD & "fsdpersonaexonerados) " _
         & "            GROUP BY cCodPers " _
         & "            HAVING SUM((nSdoDis+nIntMes)* CASE WHEN SubString(cCodCta,6,1) = '1' THEN 1 ELSE " & pnTC & " END) > " & pnFSD _
         & "        ) " _
         & "  GROUP BY cCodPers " _
         & ") A "
   
End If
dbConec.CierraConexion
If dbConec.AbreConexion Then 'Remota("07", True, False, "03")
    Set GetNroPersonas = dbConec.CargaRecordSet(sSql)
End If
dbConec.AbreConexion
End Function

Public Function GetSaldoFuncionarios(pdFecha As Date, psProd As String, psMoneda As Integer, pnTC As Currency) As Currency
Dim prs As ADODB.Recordset
sSql = "SELECT ISNULL(SUM(nSaldCnt),0) nSaldo FROM (" _
     & "   SELECT pc.cCodCta " _
     & "   FROM PersCuentaConsol pc JOIN " & sBaseFSD & "FSDPersonaExonerados pe ON pc.cCodPers = pe.cPersCod " _
     & " WHERE cRelaCta = 'TI' and cPersTipExon < 10 " _
     & " ) Cta JOIN " & sBaseFSD & "CapSaldosDiarios cs ON cs.cCodCta = Cta.cCodCta " _
     & "WHERE datediff(d,cs.dFecha,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 "

dbConec.CierraConexion
If dbConec.AbreConexion Then 'Remota("07", True, False, "03")
    Set prs = dbConec.CargaRecordSet(sSql)
    If Not prs.EOF Then
        GetSaldoFuncionarios = Round(prs!nSaldo * pnTC, 2)
    Else
        GetSaldoFuncionarios = 0
    End If
End If
dbConec.AbreConexion
End Function

Public Function CargaClientesFSD(psPersMotivo As Integer, Optional ByVal pbTodo As Boolean = False) As ADODB.Recordset
'***Parametro pbTodo agregado por ELRO el 20130507, según TI-ERS019-2013
Dim csql As String
If pnBitCentral Then
    '***Agregado por ELRO el 20130507, según TI-ERS019-2013****
    'csql = "SELECT '(' + convert(varchar(2),cPersMotivo) + ')' cTipo, p.cPersNombre cNomPers, pt.cPersCod cCodPers, cPersTipo cIdentificacion " _
    '     & "FROM " & sBaseFSD & "FSDPersonaClasif pt JOIN Persona p ON p.cPersCod = pt.cPersCod " _
    '     & "WHERE cPersMotivo = " & psPersMotivo & " ORDER BY p.cPersNombre"
    csql = "exec stp_sel_ERS0192013_CargarClientesFSD " & psPersMotivo & ", " & IIf(pbTodo, 1, 0) & ""
    '***Fin Agregado por ELRO el 20130507, según TI-ERS019-2013
Else
    csql = "SELECT '(' + convert(varchar(2),cPersMotivo) + ')' cTipo, p.cNomPers, pt.cPersCod cCodPers, cPersTipo cIdentificacion " _
         & "FROM " & sServidorFSD & "FSDPersonaClasif pt JOIN DBPERSONA..Persona p ON p.cCodPers = pt.cPersCod " _
         & "WHERE cPersMotivo = " & psPersMotivo & " ORDER BY p.cNomPers"
End If
Set CargaClientesFSD = dbConec.CargaRecordSet(csql)

End Function

Public Function GetClientesTipo2(sConUno As Boolean, sconDos As Boolean, sConTres As Boolean, sConCuatro As Boolean, sConCinco As Boolean, sConSeis As Boolean, sConSiete As Boolean, sCondElimina As String) As Integer
Dim csql As String
Dim cSqlTempo As String

On Error GoTo Error_
'***Modificado por ELRO el 20130507, según TI-ERS019-2013*****
'BLANQUEO LA TABLA Y GRABO LOS DATOS
'csql = "DELETE FROM " & sServidorFSD & "FSDPersonaClasif WHERE cPersTipo = 2 and cPersMotivo in (" & sCondElimina & ")"
'dbConec.Ejecutar (csql)
'
'csql = GetFSDSelects(2, sConUno, sconDos, sConTres, sConCuatro, sConCinco, sConSeis, sConSiete)
'csql = "SELECT DISTINCT C1.cCodPers, C1.cIdentificacion, C1.cTipo from ( " & csql & ") C1"
'dbConec.Ejecutar ("INSERT " & sServidorFSD & "FSDPersonaClasif " & csql)

csql = "exec stp_ins_ERS0192013_GenerarClientesTipo2  '" & sCondElimina & "'"
dbConec.Ejecutar (csql)

'***Fin Modificado por ELRO el 20130507, según TI-ERS019-2013
Exit Function

Error_:
Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetGenIdentClientes() As Boolean
Dim csql As String
Dim oCon As DConecta
On Error GoTo ErrorGen

'INSERTO TODOS LOS DEMAS REGISTROS
csql = " SELECT PER.cCodPers, '1' as cIdentificacion FROM dbPersona..Persona PER " & _
     " WHERE (PER.cCodPers NOT IN " & _
     " (SELECT cPersCod FROM PersFSDClasif)) "
dbConec.Ejecutar ("INSERT PersFSDClasif " & csql)
  
GetGenIdentClientes = True
Exit Function
ErrorGen:
    GetGenIdentClientes = False
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetFSDSelects(sTipoSelect As Integer, sConUno As Boolean, sconDos As Boolean, sConTres As Boolean, sConCuatro As Boolean, sConCinco As Boolean, sConSeis As Boolean, sConSiete As Boolean) As String
Dim csql As String

csql = ""
 
If sConUno = True Then
       
    csql = "SELECT "
    If pnBitCentral Then
        If sTipoSelect = 1 Then
            csql = csql & " cPersNombre cNomPers, '1' AS cTipo, "
        Else
            csql = csql & "'1' AS cTipo, "
        End If
        csql = csql & " cPersCod cCodPers, '2' AS cIdentificacion " & _
           " FROM Persona " & _
           " Where nPersPersoneria = 1 " & _
           " AND cPersNombre LIKE '%[1234567890|@#€¬$=()]%' " & _
           " Union All " & _
           " SELECT "
    
        If sTipoSelect = 1 Then
            csql = csql & " cPersNombre cNomPers, '1' AS cTipo, "
        Else
            csql = csql & "'1' AS cTipo, "
        End If
        csql = csql & " cPersCod cCodPers, '2' AS cIdentificacion " & _
           " FROM Persona " & _
           " Where nPersPersoneria <> 1 " & _
           " AND cPersNombre LIKE '%[|€¬$=]%' "
    Else
        If sTipoSelect = 1 Then
            csql = csql & " cNomPers, '1' AS cTipo, "
        Else
            csql = csql & "'1' AS cTipo, "
        End If
        csql = csql & " cCodPers, '2' AS cIdentificacion " & _
           " FROM DBPersona..Persona " & _
           " Where cTipPers = 1 " & _
           " AND CNomPers LIKE '%[1234567890|@#€¬$=()]%' " & _
           " Union All " & _
           " SELECT "
        If sTipoSelect = 1 Then
            csql = csql & " cNomPers, '1' AS cTipo, "
        Else
            csql = csql & "'1' AS cTipo, "
        End If
        csql = csql & " cCodPers, '2' AS cIdentificacion " & _
           " FROM DBPersona..Persona " & _
           " Where cTipPers <> 1 " & _
           " AND CNOMPERS LIKE '%[|€¬$=]%' "
    End If
    
End If
  
If sconDos = True Then
    If Len(Trim(csql)) > 0 Then
        csql = csql & " UNION ALL "
    End If
     
    csql = csql & " SELECT "
     
    If pnBitCentral Then
        If sTipoSelect = 1 Then
            csql = csql & " cPersNombre cNomPers, '2' AS cTipo, "
        Else
            csql = csql & "'2' AS cTipo, "
        End If
        
        csql = csql & " P.cPersCod cCodPers, '2' AS cIdentificacion " & _
                     " FROM Persona P INNER JOIN PersId PI_ ON P.cPersCod = PI_.cPersCod " & _
                     " WHERE (  LEN(PI_.cPersIdNro) < (Select PD.cPersIdSizeBeg FROM PersDocId PD WHERE PD.cPersIdTpo=PI_.cPersIdTpo) " & _
                     " OR LEN(PI_.cPersIdNro) > (Select PD.cPersIdSizeEnd FROM PersDocId PD WHERE PD.cPersIdTpo=PI_.cPersIdTpo) " & _
                     " ) "
    Else
'        If sTipoSelect = 1 Then
'            csql = csql & " cNomPers, '2' AS cTipo, "
'        Else
'            csql = csql & "'2' AS cTipo, "
'        End If
'
'        csql = csql & " P.cCodPers, '2' AS cIdentificacion " & _
'                     " FROM DBPersona..Persona P INNER JOIN PersId PI_ ON P.cCodPers = SUBSTRING(PI_.cPersCod,4,10) " & _
'                     " WHERE (  LEN(PI_.cPersIdNro) < (Select PD.cPersIdSizeBeg FROM PersDocId PD WHERE PD.cPersIdTpo=PI_.cPersIdTpo) " & _
'                     " OR LEN(PI_.cPersIdNro) > (Select PD.cPersIdSizeEnd FROM PersDocId PD WHERE PD.cPersIdTpo=PI_.cPersIdTpo) " & _
'                     " ) "
        If sTipoSelect = 1 Then
            csql = csql & " cNomPers, '2' AS cTipo, "
        Else
            csql = csql & "'2' AS cTipo, "
        End If
        
        csql = csql & " cCodPers, '2' AS cIdentificacion " & _
                     " FROM DBPersona..Persona P " & _
                     " WHERE ( " & _
                     " LEN(P.cNuDoCi) < ( " & _
                     " Select PD.cPersIdSizeBeg " & _
                     " FROM DBPersona..PersDocId PD " & _
                     " INNER JOIN dbComunes..TablaCod TC on " & _
                     " TC.cCodTab = PD.cPersIdtpo " & _
                     " WHERE PD.cPersIdtpo like '04%' " & _
                     " AND TC.cValor = P.cTiDoCi ) " & _
                     " OR LEN(P.cNuDoCi) > ( " & _
                     " Select PD.cPersIdSizeEnd " & _
                     " FROM DBPersona..PersDocId PD " & _
                     " INNER JOIN dbComunes..TablaCod TC on " & _
                     " TC.cCodTab = PD.cPersIdtpo " & _
                     " WHERE PD.cPersIdtpo like '04%' " & _
                     " AND TC.cValor = P.cTiDoCi ) " & _
                     " OR LEN(P.cNuDoTr) < ( " & _
                     " Select PD.cPersIdSizeBeg " & _
                     " FROM DBPersona..PersDocId PD " & _
                     " INNER JOIN dbComunes..TablaCod TC on " & _
                     " TC.cCodTab = PD.cPersIdtpo " & _
                     " WHERE PD.cPersIdtpo like '05%' " & _
                     " AND TC.cValor = P.cTiDoTr ) " & _
                     " OR LEN(P.cNuDoTr) > ( "
        csql = csql & " Select PD.cPersIdSizeEnd " & _
                     " FROM DBPersona..PersDocId PD " & _
                     " INNER JOIN dbComunes..TablaCod TC on " & _
                     " TC.cCodTab = PD.cPersIdtpo " & _
                     " WHERE PD.cPersIdtpo like '05%' " & _
                     " AND TC.cValor = P.cTiDoTr  ))"
    End If
End If
  
If sConTres = True Then
    If Len(Trim(csql)) > 0 Then
        csql = csql & " UNION ALL "
    End If
    
    csql = csql & " SELECT "
     
    If pnBitCentral Then
        If sTipoSelect = 1 Then
            csql = csql & " cPersNombre cNomPers, '3' AS cTipo, "
        Else
            csql = csql & "'3' AS cTipo, "
        End If
        
        csql = csql & " P.cPersCod cCodPers, '2' AS cIdentificacion " & _
                      " FROM Persona P INNER JOIN PersId PI_ ON P.cPersCod = PI_.cPersCod " & _
                      " WHERE P.nPersPersoneria = '" & gPersonaNat & "' " & _
                      " AND PI_.CPersIdNro LIKE '%[|@#€¬$=()]%' " & _
                      " Union All " & _
                      " Select "
                      
        If sTipoSelect = 1 Then
            csql = csql & " cPersNombre cNomPers, '3' AS cTipo, "
        Else
            csql = csql & "'3' AS cTipo, "
        End If
        
        csql = csql & " P.cPersCod cCodPers, '2' AS cIdentificacion " & _
                      " FROM Persona P INNER JOIN PersId PI_ ON P.cPersCod = PI_.cPersCod " & _
                      " WHERE P.nPersPersoneria = '" & gPersonaNat & "' " & _
                      " AND PI_.CPersIdNro LIKE '%[|€¬$=]%' "
    Else
'        If sTipoSelect = 1 Then
'            csql = csql & " cNomPers, '3' AS cTipo, "
'        Else
'            csql = csql & "'3' AS cTipo, "
'        End If
'
'        csql = csql & " cCodPers, '2' AS cIdentificacion " & _
'                      " FROM DBPersona..Persona P INNER JOIN PersId PI_ ON P.cCodPers = SUBSTRING(PI_.cPersCod,4,10) " & _
'                      " WHERE P.CTipPers='1' " & _
'                      " AND PI_.CPersIdNro LIKE '%[|@#€¬$=()]%' " & _
'                      " Union All " & _
'                      " Select "
'
'        If sTipoSelect = 1 Then
'            csql = csql & " cNomPers, '3' AS cTipo, "
'        Else
'            csql = csql & "'3' AS cTipo, "
'        End If
'
'        csql = csql & " cCodPers, '2' AS cIdentificacion " & _
'                      " FROM DBPersona..Persona P INNER JOIN PersId PI_ ON P.cCodPers = SUBSTRING(PI_.cPersCod,4,10) " & _
'                      " WHERE P.CTipPers<>'1' " & _
'                      " AND PI_.CPersIdNro LIKE '%[|€¬$=]%' "
        If sTipoSelect = 1 Then
            csql = csql & " cNomPers, '3' AS cTipo, "
        Else
            csql = csql & "'3' AS cTipo, "
        End If
        
        csql = csql & " cCodPers, '2' AS cIdentificacion " & _
                      " FROM DBPersona..Persona P " & _
                      " Inner Join " & _
                      " DBPersona..PersDocId PD " & _
                      " on P.cTidoci=1 AND P.cNudoci LIKE '%[|@#€¬$=()]%' " & _
                      " Union All " & _
                      " Select "
                      
        If sTipoSelect = 1 Then
            csql = csql & " cNomPers, '3' AS cTipo, "
        Else
            csql = csql & "'3' AS cTipo, "
        End If
        
        csql = csql & " cCodPers, '2' AS cIdentificacion " & _
                      " FROM DBPersona..Persona P " & _
                      " Inner Join " & _
                      " DBPersona..PersDocId PD " & _
                      " on P.cTidoci<>1 AND P.cNudoci LIKE '%[|€¬$=]%' "
    End If
End If
 
If sConCuatro = True Then
    If Len(Trim(csql)) > 0 Then
        csql = csql & " UNION ALL "
    End If
    
    csql = csql & " SELECT "
     
    If pnBitCentral Then
        If sTipoSelect = 1 Then
            csql = csql & " c1.cPersNombre cNomPers, '4' AS cTipo, "
        Else
            csql = csql & "'4' AS cTipo, "
        End If
        
        csql = csql & " c2.cPersCod cCodPers, '2' as cIdentificacion " & _
                      " From " & _
                      " ( " & _
                      " SELECT cPersNombre " & _
                      " FROM Persona P INNER JOIN  " & _
                      "      PersID ON P.cPersCod = dbo.PersID.cPersCod " & _
                      " GROUP BY  P.cPersNombre, PersID.cPersIDTpo, PersID.cPersIDnro " & _
                      " Having Count(cPersNombre) > 1 " & _
                      " ) C1 " & _
                      " Inner Join  " & _
                      " ( " & _
                      " Select P.cPersCod, P.cPersNombre FROM Persona P " & _
                      " ) C2 " & _
                      " ON C1.cPersNombre =C2.cPersNombre "
    Else
'        If sTipoSelect = 1 Then
'            csql = csql & " c1.cNomPers, '4' AS cTipo, "
'        Else
'            csql = csql & "'4' AS cTipo, "
'        End If
'
'        csql = csql & " c2.cCodPers, '2' as cIdentificacion " & _
'                      " From " & _
'                      " ( " & _
'                      " SELECT  cNomPers " & _
'                      " FROM dbPersona..Persona P INNER JOIN  " & _
'                      " dbo.PersID ON P.cCodPers = SUBSTRING(dbo.PersID.cPersCod, 4, 10) " & _
'                      " GROUP BY  P.cNomPers, PersID.cPersIDTpo, PersID.cPersIDnro " & _
'                      " Having Count(cNomPers) > 1 " & _
'                      " ) C1 " & _
'                      " Inner Join  " & _
'                      " ( " & _
'                      " Select P.cCodPers, P.cNomPers FROM DBPersona..Persona P " & _
'                      " ) C2 " & _
'                      " ON C1.cNomPers=C2.cNomPers "
        If sTipoSelect = 1 Then
            csql = csql & " c1.cNomPers, '4' AS cTipo, "
        Else
            csql = csql & "'4' AS cTipo, "
        End If
        
        csql = csql & " c2.cCodPers, '2' as cIdentificacion " & _
                      " From ( " & _
                      " select distinct  cnompers from ( " & _
                      " SELECT  cNomPers " & _
                      " FROM dbPersona..Persona P " & _
                      " GROUP BY  P.cNomPers, P.cTiDoCi, P.cNuDoCi " & _
                      " Having Count(cNomPers) > 1 " & _
                      " Union " & _
                      " SELECT  cNomPers " & _
                      " FROM dbPersona..Persona P " & _
                      " GROUP BY  P.cNomPers, P.cTiDoTr, P.cNuDoTr " & _
                      " Having Count(cNomPers) > 1 " & _
                      " ) SubC "
        csql = csql & " ) C1 " & _
                      " Inner Join ( " & _
                      " Select P.cCodPers, P.cNomPers FROM DBPersona..Persona P  " & _
                      " ) C2 " & _
                      " ON C1.cNomPers=C2.cNomPers "
    End If
End If
 
If sConCinco = True Then
    If Len(Trim(csql)) > 0 Then
        csql = csql & " UNION ALL "
    End If
     
    csql = csql & " SELECT "
     
    If pnBitCentral Then
        If sTipoSelect = 1 Then
            csql = csql & " P.cPersNombre cNomPers, '5' AS cTipo, "
        Else
            csql = csql & "'5' AS cTipo, "
        End If
        
        csql = csql & " P.cPersCod cCodPers, '2' AS cIdentificacion FROM " & _
                      " PERSONA P " & _
                      " GROUP BY P.cPersCod, P.cPersNombre " & _
                      " Having Count(P.cPersCod) > 1 "
    Else
'        If sTipoSelect = 1 Then
'            csql = csql & " P.cNomPers, '5' AS cTipo, "
'        Else
'            csql = csql & "'5' AS cTipo, "
'        End If
'
'        csql = csql & " P.cCodPers, '2' AS cIdentificacion FROM " & _
'                      " DBPERSONA..PERSONA P " & _
'                      " GROUP BY P.cCodPers, P.cNomPers " & _
'                      " Having Count(P.cCodPers) > 1 "

        If sTipoSelect = 1 Then
            csql = csql & " P.cNomPers, '5' AS cTipo, "
        Else
            csql = csql & "'5' AS cTipo, "
        End If
        
        csql = csql & " P.cCodPers, '2' AS cIdentificacion FROM " & _
                      " DBPERSONA..PERSONA P " & _
                      " GROUP BY P.cCodPers, P.cNomPers " & _
                      " Having Count(P.cCodPers) > 1 "
    End If
End If
 
If sConSeis = True Then
    If Len(Trim(csql)) > 0 Then
        csql = csql & " UNION ALL "
    End If
    csql = csql & " SELECT "
     
    If pnBitCentral Then
        If sTipoSelect = 1 Then
            csql = csql & " C2.cPersNombre cNomPers, '6' AS cTipo, "
        Else
            csql = csql & "'6' AS cTipo, "
        End If
        
        csql = csql & " C2.cCodPers, '2' AS cIdentificacion FROM " & _
                      " ( " & _
                      " SELECT PersID.cPersIDTpo AS cTipo, PersID.cPersIDnro as cNro " & _
                      " FROM Persona P INNER JOIN " & _
                      "    PersID ON P.cPersCod = dbo.PersID.cPersCod " & _
                      " GROUP BY PersID.cPersIDTpo, PersID.cPersIDnro " & _
                      " Having (Count(dbo.PersID.cPersIDnro) > 1) And (Count(PersID.cPersIDTpo) > 1) " & _
                      " ) C1 " & _
                      " Inner Join " & _
                      " ( " & _
                      " Select P.cPersCod cCodPers, cPersNombre cNomPers, PI_.cPersIdTpo as cTipo, PI_.cPersIdNro as cNro " & _
                      " FROM Persona P INNER JOIN PersId PI_ ON P.cPersCod = PI_.cPersCod " & _
                      " ) C2 " & _
                      "ON C1.cTipo=C2.cTipo AND c1.cNro=C2.cNro "
    Else
        If sTipoSelect = 1 Then
            csql = csql & " C2.cNomPers, '6' AS cTipo, "
        Else
            csql = csql & "'6' AS cTipo, "
        End If
        
        csql = csql & " C2.cCodPers, '2' AS cIdentificacion FROM " & _
                      " ( Select distinct aa.cTipo, aa.cNro from " & _
                      " ( SELECT P.cTiDoCi AS cTipo, P.cNuDoCi as cNro " & _
                      " FROM dbPersona..Persona P WHERE not RTRIM(P.cNuDoCi) = '' " & _
                      " GROUP BY P.cTiDoCi, P.cNuDoCi " & _
                      " Having (Count(P.cNuDoCi) > 1) And (Count(P.cTiDoCi) > 1) " & _
                      " Union " & _
                      " SELECT P.cTiDoTr AS cTipo, P.cNuDoTr as cNro " & _
                      " FROM dbPersona..Persona P WHERE not rTRIM(P.cNuDoTr) = '' " & _
                      " GROUP BY P.cTiDoTr, P.cNuDoTr " & _
                      " Having (Count(P.cNuDoTr) > 1) And (Count(P.cTiDoTr) > 1) " & _
                      " ) aa ) C1 " & _
                      " Inner Join " & _
                      " ( "
            csql = csql & "select distinct bb.cCodPers, bb.cNomPers, bb.cTipo, bb.cNro " & _
                      " from ( " & _
                      " Select P.cCodPers, P.cNomPers, P.cTiDoCi as cTipo, P.cNuDoCi as cNro " & _
                      " FROM DBPersona..Persona P " & _
                      " Union " & _
                      " Select P.cCodPers, P.cNomPers, P.cTiDoTr as cTipo, P.cNuDoTr as cNro " & _
                      " FROM DBPersona..Persona P ) bb " & _
                      " ) C2 ON C1.cTipo=C2.cTipo AND c1.cNro=C2.cNro "
    End If
End If
 
If sConSiete = True Then
    If Len(Trim(csql)) > 0 Then
        csql = csql & " UNION ALL "
    End If
    
    csql = csql & " SELECT "
     
    If pnBitCentral Then
        If sTipoSelect = 1 Then
            csql = csql & " cPersNombre cNomPers, '7' AS cTipo, "
        Else
            csql = csql & "'7' AS cTipo, "
        End If
        
        csql = csql & " cPersCod cCodPers, '2' AS cIdentificacion " & _
                      " FROM PERSONA " & _
                      " Where nPersPersoneria = 1 " & _
                      " AND ( " & _
                      "    cPersNombre LIKE '%S.A%' " & _
                      " OR cPersNombre LIKE '%S.R%' " & _
                      " OR cPersNombre LIKE '%S.C%' " & _
                      " OR cPersNombre LIKE '%LTD%' " & _
                      " ) "
    Else
'        If sTipoSelect = 1 Then
'            csql = csql & " cNomPers, '7' AS cTipo, "
'        Else
'            csql = csql & "'7' AS cTipo, "
'        End If
'
'        csql = csql & " cCodPers, '2' AS cIdentificacion " & _
'                      " FROM DBPERSONA..PERSONA " & _
'                      " Where CTIPPERS = 1 " & _
'                      " AND ( " & _
'                      " CNOMPERS LIKE '%S.A%' " & _
'                      " OR CNOMPERS LIKE '%S.R%' " & _
'                      " OR CNOMPERS LIKE '%S.C%' " & _
'                      " OR CNOMPERS LIKE '%LTD%' " & _
'                      " ) "

    If sTipoSelect = 1 Then
        csql = csql & " cNomPers, '7' AS cTipo, "
    Else
        csql = csql & "'7' AS cTipo, "
    End If
    
    csql = csql & " cCodPers, '2' AS cIdentificacion " & _
                  " FROM DBPERSONA..PERSONA " & _
                  " Where CTIPPERS = 1 " & _
                  " AND ( " & _
                  " CNOMPERS LIKE '%S.A%' " & _
                  " OR CNOMPERS LIKE '%S.R%' " & _
                  " OR CNOMPERS LIKE '%S.C%' " & _
                  " OR CNOMPERS LIKE '%LTD%' " & _
                  " ) "
    End If
End If

GetFSDSelects = csql ' & IIf(sTipoSelect = 1, " ORDER BY cTipo, cCodPers", " ")

End Function
 
''''''''''''''''''''
'Clientes Exonerados
''''''''''''''''''''
Public Sub InsertaNuevoExonerado(ByVal psPersCod As String, ByVal pnTipoExo As Integer, ByVal psNotas As String, ByVal psMovNro As String)
'***Parametro psMovNro agregado y parametro pnTipoExo modificado  por ELRO el 20130507, según TI-ERS019-2013
Dim sSql As String

On Error GoTo ErrorInsertaNuevoExonerado
    '***Modificado por ELRO el 20130507, según TI-ERS019-2013****
    'sSql = "INSERT INTO " & sServidorFSD & "FSDPersonaExonerados (cPersCod,cPersTipExon,cPersNotas)"
    'sSql = sSql & "VALUES('" & Trim(psPersCod) & "'," & psTipoExo & ",'" & psNotas & "')"
    sSql = "exec stp_ins_ERS0192013_RegistrarExonerado '" & psPersCod & "', " & pnTipoExo & ", '" & psNotas & "', '" & psMovNro & "'"
    '***Fin Modificado por ELRO el 20130507, según TI-ERS019-2013
    dbConec.Ejecutar sSql
 
    Exit Sub
ErrorInsertaNuevoExonerado:
    Err.Raise Err.Number, "Nuevo Exonerado", Err.Description
End Sub

Public Sub EliminaExonerado(ByVal psPersCod As String, ByVal psMovNro As String)
'***Parametro psMovNro agregado por ELRO el 20130507, según TI-ERS019-2013
Dim sSql As String

On Error GoTo ErrorEliminaExonerado
    '***Modificado por ELRO el 20130507, según TI-ERS019-2013****
    'sSql = "DELETE FROM " & sServidorFSD & "FSDPersonaExonerados  WHERE cPersCod='" & psPersCod & "'"
    sSql = "exec stp_upd_ERS0192013_EliminarExonerado  '" & psPersCod & "', '" & psMovNro & "'"
    '***Fin Modificado por ELRO el 20130507, según TI-ERS019-2013
    dbConec.Ejecutar sSql
 
    Exit Sub
ErrorEliminaExonerado:
    Err.Raise Err.Number, "Elimina Exonerado", Err.Description
End Sub

Public Function GetExisteExonerado(ByVal psPersCod As String) As Boolean
Dim sSql As String
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset

On Error GoTo ErrorGetExisteExonerado
    '***Modificado por ELRO el 20130507, según TI-ERS019-2013***
    'sSql = "select * FROM " & sServidorFSD & "FSDPersonaExonerados  WHERE cPersCod=" & psPersCod
    sSql = "exec stp_sel_ERS0192013_VerificarExonerado " & psPersCod
    '***Fin Modificado por ELRO el 20130507, según TI-ERS019-2013
    Set rs = dbConec.CargaRecordSet(sSql)
    If rs.BOF Then
        GetExisteExonerado = False
    Else
        GetExisteExonerado = True
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrorGetExisteExonerado:
    Err.Raise Err.Number, "Verificar existe Exonerado", Err.Description
End Function

Public Function GetListaExonerados() As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim csql As String
Set rs = New ADODB.Recordset
On Error GoTo ErrorGetListaExonerados
If pnBitCentral Then
    '***Modificado por ELRO el 20130507, según TI-ERS019-2013***
    'csql = " SELECT PER.cPersCod as cCOdigoMostrar, PFSD.cPersCod as cCodPers, Per.cPersNombre, CON.cConsDescripcion as cPersTipExon, PFSD.cPersNotas " & _
    '        " FROM " & sServidorFSD & "FSDPersonaExonerados  PFSD INNER JOIN " & _
    '        " Persona PER ON PFSD.cPersCod = PER.cPersCod " & _
    '        " INNER JOIN CONSTANTE CON ON Con.nConsCod LIKE '" & gFSDTipoExonerados & "' AND Con.nConsValor=PFSD.cPersTipExon " & _
    '        " ORDER BY PFSD.cPersTipExon, cPersNombre "
    csql = "exec stp_sel_ERS0192013_ListarExonerados "
    '***Fin Modificado por ELRO el 20130507, según TI-ERS019-2013
Else
'    csql = " SELECT PER.cPersCod as cCOdigoMostrar, PFSD.cPersCod as cCodPers, Per.cPersNombre, CON.cConsDescripcion as cPersTipExon, PFSD.cPersNotas " & _
'            " FROM " & sServidorFSD & "FSDPersonaExonerados  PFSD INNER JOIN " & _
'            " Persona PER ON PFSD.cPersCod = right(PER.cPersCod,10) " & _
'            " INNER JOIN CONSTANTE CON ON Con.nConsCod LIKE '" & gFSDTipoExonerados & "' AND Con.nConsValor=PFSD.cPersTipExon " & _
'            " ORDER BY PFSD.cPersTipExon, cPersNombre "
            
    csql = " SELECT Per.cPersCod as cCodigoMostrar, PFSD.cPersCod as cCodPers, "
    csql = csql & " Per.cPersNOmbre as cPersNombre, CON.cConsDescripcion as cPersTipExon, PFSD.cPersNotas "
    csql = csql & " FROM " & sServidorFSD & "FSDPersonaExonerados  PFSD "
    csql = csql & " INNER JOIN persona PER ON PFSD.cPersCod = PER.cPersCod "
    csql = csql & " INNER JOIN CONSTANTE CON ON Con.nConsCod LIKE '" & gFSDTipoExonerados & "' AND Con.nConsValor=PFSD.cPersTipExon "
    csql = csql & " ORDER BY PFSD.cPersTipExon, cPersNombre "
End If
Set rs = dbConec.CargaRecordSet(csql)
dbConec.AbreConexion
Set GetListaExonerados = rs
Exit Function

ErrorGetListaExonerados:
    Err.Raise Err.Number, "Lista Exonerados", Err.Description
    
End Function
 
 
Private Sub Class_Initialize()
Dim oImp As DImpresoras
Dim oConst As NConstSistemas
Dim oIni As ClasIni
 
Set oImp = New DImpresoras
oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
Set oImp = Nothing

Set oIni = New ClasIni
vsBaseComunes = oIni.BaseComunes
vsBasePersonas = oIni.BasePersonas
Set oIni = Nothing

sLPT = "LPT1"
BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")

Set oConst = New NConstSistemas
pnBitCentral = IIf(oConst.LeeConstSistema(gConstSistBitCentral) = "1", True, False)
sServidorFSD = oConst.LeeConstSistema(gConstSistServFSD)
sBaseFSD = sServidorFSD  ' Mid(sServidorFSD, InStr(sServidorFSD, "]") + 2)
'sServidorConsolidada = oConst.LeeConstSistema(gConstSistServCentralRiesgos)
Set oConst = Nothing

Set dbConec = New DConecta
dbConec.AbreConexion 'Remota gsCodAge, True, False, "03"
'dbConec.CommandTimeOut (9000)
End Sub

Private Sub Class_Terminate()
dbConec.CierraConexion
Set dbConec = Nothing
End Sub
