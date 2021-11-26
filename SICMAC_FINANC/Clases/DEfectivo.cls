VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Defectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsServerCom As String
Dim vsServerPers As String
Dim vsFormatoFecha As String
Public Sub Inicio(ByVal psFormatoFecha As String)
vsFormatoFecha = psFormatoFecha
End Sub
Public Function EmiteBilletajes(ByVal psMonedas As Moneda, Optional ByVal psDenominacion As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsFiltroDenom As String
Set oConect = New DConecta
If psDenominacion <> "" Then
    lsFiltroDenom = " and Substring(cEfectivoCod,2,1)='" & psDenominacion & "' "
End If
If oConect.AbreConexion = False Then Exit Function

sql = " Select     Case " _
    & "         WHEN Substring(cEfectivoCod,2,1)='B'  Then 'BILLETAJE' + Convert(char(10),nEfectivoValor) " _
    & "         WHEN Substring(cEfectivoCod,2,1)='M'  Then 'MONEDA   ' + Convert(char(10),nEfectivoValor) " _
    & "         END as Descripcion, 0 as Cantidad, 0  as Monto, cEfectivoCod, nEfectivoValor " _
    & " From Efectivo " _
    & " WHERE SUBSTRING(cEfectivoCod,1,1) ='" & psMonedas & "' " & lsFiltroDenom _
    & " ORDER BY  SUBSTRING(cEfectivoCod,2,1), nEfectivoValor DESC "
    
Set rs = oConect.CargaRecordSet(sql)
Set EmiteBilletajes = rs

oConect.CierraConexion: Set oConect = Nothing
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

End Sub
Public Function GetBilletajes(ByVal psFiltro As String) As ADODB.Recordset
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltro As String

Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
If psFiltro <> "" Then
    lsFiltro = " WHERE cEfecTivoCod like '" & psFiltro & "%'"
End If

sql = "   SELECT  cEfectivoCod, " _
    & "         CASE " _
    & "         WHEN SUBSTRING(cEfectivoCod,2,1)='B' THEN 'BILLETES ' + Convert(Char(10),nEfectivoValor) " _
    & "         WHEN SUBSTRING(cEfectivoCod,2,1)='M' THEN 'MONEDAS  ' + Convert(Char(10),nEfectivoValor) " _
    & "         END AS Descripcion , 1  as Nivel " _
    & "  From " & vsServerCom & "EFECTIVO " & lsFiltro _
    & "  ORDER BY SUBSTRING(cEfectivoCod,2,1), nEfectivoValor DESC "

Set rs = oConect.CargaRecordSet(sql)
Set GetBilletajes = rs
oConect.CierraConexion: Set oConect = Nothing

End Function
Public Function GetBilletajesMov(ByVal psMovNro As String, Optional ByVal pnMoneda As Moneda, Optional ByVal psDenominacion As String = "")
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltroDenom As String

Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

If psDenominacion <> "" Then
    lsFiltroDenom = " and Substring(E.cEfectivoCod,2,1)='" & psDenominacion & "' "
End If

sql = "  SELECT Case " _
    & "             WHEN Substring(E.cEfectivoCod,2,1)='B'  Then 'BILLETAJE' + Convert(char(10),E.nEfectivoValor) " _
    & "             WHEN Substring(E.cEfectivoCod,2,1)='M'  Then 'MONEDA   ' + Convert(char(10),E.nEfectivoValor) " _
    & "         END as Descripcion ,  ISNULL(MovEfect.Cantidad,0) as Cantidad , ISNULL(MovEfect.Importe,0) as Monto, " _
    & "         E.cEfectivoCod , E.nEfectivoValor " _
    & "  FROM   EFECTIVO E " _
    & "         Left Join " _
    & "         (SELECT MOE.nMonto/E.nEfectivoValor  AS CANTIDAD, " _
    & "                 MOE.nMonto AS IMPORTE, MOE.cEfectivoCod, E.nEfectivoValor , M.cOpeCod " _
    & "          FROM   MOV M " _
    & "                 JOIN MOVUSEREFECTIVO MOE ON MOE.nMovNro =M.nMovNro " _
    & "                 JOIN EFECTIVO E ON E.cEfectivoCod = MOE.cEfectivoCod " _
    & "          WHERE   M.CMOVNRO= '" & psMovNro & "') AS MovEfect " _
    & "     ON MovEfect.cEfectivoCod = E.cEfectivoCod " _
    & " WHERE E.cEfectivoCod LIKE '" & pnMoneda & "%' " & lsFiltroDenom & " Order BY E.nEfectivoValor Desc "

Set rs = oConect.CargaRecordSet(sql)
Set GetBilletajesMov = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetEfectivoMovImp(ByVal psMovNro As String, Optional ByVal pbDebe As Boolean = True) As ADODB.Recordset
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltro As String

Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

sql = "SELECT  CMOVNRO, " _
    & "         CASE WHEN Tipo ='B' THEN nEfectivoValor  ELSE 0 END AS Billetes, " _
    & "         SUM(CASE WHEN Tipo = 'B' THEN ABS(Importe) ELSE 0 END ) AS MontoBilletes, " _
    & "         CASE WHEN Tipo ='M' THEN nEfectivoValor  ELSE 0 END AS Monedas, " _
    & "         SUM(CASE WHEN Tipo = 'M' THEN ABS(Importe) ELSE 0 END ) AS MontoMonedas , Moneda " _
    & " FROM (SELECT    M.CMOVNRO, E.nEfectivoValor,Substring(e.cefectivocod,2,1) AS  Tipo, Substring(e.cefectivocod,1,1) AS Moneda ,  " _
    & "                 ISNULL(Me.nMovMEImporte, MC.nMovImporte) As Importe " _
    & "        FROM     MOVOBJEFECTIVO MOE JOIN EFECTIVO E ON E.CEFECTIVOCOD = MOE.CEFECTIVOCOD  " _
    & "                 JOIN MOVCTA MC ON MC.NMOVNRO = MOE.NMOVNRO AND MC.NMOVITEM = MOE.NMOVITEM JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO  " _
    & "                 LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO  AND ME.NMOVITEM = MC.NMOVITEM " _
    & "         WHERE  M.CMOVNRO ='" & psMovNro & "'  AND MC.NMOVIMPORTE" & IIf(pbDebe, ">", "<") & "0  ) AS MC " _
    & " GROUP BY CMOVNRO , nEfectivoValor , Tipo , Moneda " _
    & " ORDER BY nEfectivoValor "

Set rs = oConect.CargaRecordSet(sql)
Set GetEfectivoMovImp = rs
oConect.CierraConexion: Set oConect = Nothing

End Function

Public Function BilletajeCajaAgencias(psEfectivoCod As String, lsMoneda As String, ldFecha As Date, lsCodAge As String) As Currency
Dim SQLBil  As String
Dim lsFiltro As String
Dim rs      As New ADODB.Recordset
Dim oConect As New DConecta
   
Dim SQLFil As String 'Agregado por PASI20131029 segun INC1310230010
Dim ldFechaFiltrada As Date 'Agregado por PASI20131029 segun INC1310230010
Dim rsFiltro As New ADODB.Recordset 'Agregado por PASI20131029 segun INC1310230010
Dim oConectFiltro As New DConecta 'Agregado por PASI20131029 segun INC1310230010
Dim bExisteSaldo As Boolean
    
    'Agregado por PASI20131029 segun INC1310230010
    ldFechaFiltrada = ldFecha
    
    '*****Comentado x PASI2010509
    'oConectFiltro.AbreConexion
    'SQLFil = "select dbo.fnc_EsFeriadoDomingo('" & Format(ldFechaFiltrada, gsFormatoMovFecha) & "', '" & lsCodAge & "') as Resultado"
    'Set rsFiltro = oConectFiltro.CargaRecordSet(SQLFil)
    'If Not rsFiltro.EOF Then
    '    Do While rsFiltro!Resultado > 0
    '        RSClose rsFiltro
    '        ldFechaFiltrada = DateAdd("d", -1, ldFechaFiltrada)
    '        SQLFil = "select dbo.fnc_EsFeriadoDomingo('" & Format(ldFechaFiltrada, gsFormatoMovFecha) & "', '" & lsCodAge & "') as Resultado"
    '        Set rsFiltro = oConectFiltro.CargaRecordSet(SQLFil)
    '    Loop
    'End If
    '****************************
    
    'Fin Agregado por PASI20131029 segun INC1310230010
   
   
   oConect.AbreConexion
   BilletajeCajaAgencias = 0
   If psEfectivoCod <> "" Then
        lsFiltro = " and Mue.cEfectivoCod = '" & psEfectivoCod & "' "
   End If
   
   'SQLBil = "SELECT ISNULL(SUM(nMonto),0) nMonto " _
         & "FROM   MovUserEfectivo MUE JOIN Mov M ON MUE.nMovNro = M.nMovNro " _
         & "WHERE  m.cMovNro = (SELECT Max(cMovNro) FROM Mov m1 join MovUserEfectivo mu ON mu.nMovNro = m1.nMovNro WHERE M1.nMovEstado = " & gMovEstContabNoContable & " and M1.nMovFlag = " & gMovFlagVigente & " and LEFT(m1.cMovNro,8) <= '" & Format(ldFecha, gsFormatoMovFecha) & "' and " _
         & "                    LEFT(mu.cEfectivoCod,1) = " & lsMoneda & " and m1.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') and " _
         & "                    SubString(m1.cMovNro,18,2) = '" & lsCodAge & "') " & lsFiltro
   
   'Comentado por PASI20131029 segun INC1310230010
        'SQLBil = " select ISNULL(SUM(MUE.nMonto),0) nMonto  " _
        '    & "     from    movuserefectivo MUE " _
        '    & "             join mov m on m.nmovnro = MUE.nmovnro " _
        '    & "     where   substring(m.cmovnro,18,2)='" & lsCodAge & "' and m.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') " _
        '    & "             and left(m.cmovnro,8) = '" & Format(ldFecha, gsFormatoMovFecha) & "' and m.nmovflag = 0 and LEFT(MUE.cEfectivoCod,1) = " & lsMoneda & " " & lsFiltro
        
    'Agregado por PASI20131029 segun INC1310230010
    'Comentado PASI20160509**********
'    SQLBil = " select ISNULL(SUM(MUE.nMonto),0) nMonto  " _
'        & "     from    movuserefectivo MUE " _
'        & "             join mov m on m.nmovnro = MUE.nmovnro " _
'        & "     where   convert(char(2),substring(m.cmovnro,18,2))='" & lsCodAge & "' and m.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') " _
'        & "             and convert(date,left(m.cmovnro,8)) = '" & Format(ldFechaFiltrada, gsFormatoMovFecha) & "' and m.nmovflag = 0 and convert(integer,LEFT(MUE.cEfectivoCod,1)) = " & lsMoneda & " " & lsFiltro
    'end PASI************************
    'Fin Agregado por PASI20131029 segun INC1310230010
    
    'PASI20160509************
    If Not AgenciaEsActiva(lsCodAge) = 0 Then
        bExisteSaldo = False
               
        Do While Not bExisteSaldo
            oConectFiltro.AbreConexion
            SQLFil = "select dbo.fnc_EsFeriadoDomingo('" & Format(ldFechaFiltrada, gsFormatoMovFecha) & "', '" & lsCodAge & "') as Resultado"
            Set rsFiltro = oConectFiltro.CargaRecordSet(SQLFil)
        
            SQLBil = " select ISNULL(SUM(MUE.nMonto),0) nMonto  " _
            & "     from    movuserefectivo MUE " _
            & "             join mov m on m.nmovnro = MUE.nmovnro " _
            & "     where   convert(char(2),substring(m.cmovnro,18,2))='" & lsCodAge & "' and m.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') " _
            & "             and convert(date,left(m.cmovnro,8)) = '" & Format(ldFechaFiltrada, gsFormatoMovFecha) & "' and m.nmovflag = 0 and convert(integer,LEFT(MUE.cEfectivoCod,1)) = " & lsMoneda & " " & lsFiltro
            Set rs = oConect.CargaRecordSet(SQLBil)
            If Not RSVacio(rs) Then
                If rs!nMonto = 0 And rsFiltro!Resultado > 0 Then
                    ldFechaFiltrada = DateAdd("d", -1, ldFechaFiltrada)
                Else
                    BilletajeCajaAgencias = rs!nMonto
                    bExisteSaldo = True
                End If
            End If
        Loop
    Else
        BilletajeCajaAgencias = 0
    End If
    'end PASI****************
    
'Comentado xPASI20160509********
'   Set rs = oConect.CargaRecordSet(SQLBil)
'   If Not RSVacio(rs) Then
'      BilletajeCajaAgencias = rs!nMonto
'   End If
'end **************************

   RSClose rs
   oConect.CierraConexion
   Set oConect = Nothing
End Function


Public Function Operacion_FlujoDiariaCaja(psCtaCaja As String, pdFecha As Date, psOpecod As String, pbIngreso As Boolean) As ADODB.Recordset
Dim sql As String
Dim oConect As DConecta

sql = "SELECT ot.cOpeCod, ot.cOpeDesc, oc.cOpeCtaDH, " _
     & "       SUM(mc.nMovImporte * " & IIf(pbIngreso, "1", "-1") & ") nMovImporte, ISNULL(SUM(me.nMovMEImporte * " & IIf(pbIngreso, "1", "-1") & "),0) nMovMEImporte " _
     & "FROM Mov m  " _
     & "     JOIN MovCta mc ON mc.nMovNro = m.nMovNro " _
     & "     LEFT JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
     & "     JOIN OpeCta oc ON oc.cOpecod = m.cOpeCod, OpeTpo Ot " _
     & "WHERE m.cOpeCod LIKE '4%' and m.nMovEstado = '" & gMovEstContabMovContable & "' and m.nMovFlag in ('" & gMovFlagVigente & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "')" _
     & "     and m.cMovNro LIKE '" & Format(pdFecha, gsFormatoMovFecha) & "%' and oc.cCtaContCod = '" & psCtaCaja & "' " _
     & "     and mc.cCtaContCod = '" & psCtaCaja & "' and cOpeCtaDH = '" & IIf(pbIngreso, "D", "H") & "' and mc.nMovImporte " & IIf(pbIngreso, ">", "<") & " 0 " _
     & "     and ot.cOpeCod = ( SELECT MAX(cOpeCod) FROM OpeTpo " _
     & "                        WHERE cOpeCod <= m.cOpeCod and nOpeNiv = 1 ) " _
     & "GROUP BY ot.cOpeCod, ot.cOpeDesc, oc.cOpeCtaDH"

Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
     
Set Operacion_FlujoDiariaCaja = oConect.CargaRecordSet(sql)
oConect.CierraConexion: Set oConect = Nothing

End Function
'PASI20160509*****
Public Function AgenciaEsActiva(ByVal psAgeCod As String) As Integer
Dim rs      As New ADODB.Recordset
Dim oConect As New DConecta
Dim sSQL As String
    oConect.AbreConexion
    sSQL = "select nEstado From Agencias where cagecod = '" & psAgeCod & "'"
    Set rs = oConect.CargaRecordSet(sSQL)
    If Not RSVacio(rs) Then
        AgenciaEsActiva = rs!nEstado
    Else
        AgenciaEsActiva = 0
    End If
End Function
'end PASI******
'********************NAGL 20180920 Según TIC1807210002**************
Public Function EmiteBilletajeMonedaNew(pdFecha As Date, psMoneda As String, psCtaContCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConect As New DConecta
    oConect.AbreConexion
    sSQL = "Exec stp_sel_GetEmiteBilletajeMonDescomp '" & Format(pdFecha, "yyyymmdd") & "','" & psMoneda & "','" & psCtaContCod & "'"
    Set EmiteBilletajeMonedaNew = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Public Function ObtieneTipoEfectivo(psMoneda As String) As ADODB.Recordset
Dim sSQL As String
Dim oConect As New DConecta
    oConect.AbreConexion
    sSQL = "Exec stp_Sel_GetTipoEfectivo '" & psMoneda & "'"
    Set ObtieneTipoEfectivo = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Public Function ObtieneIndiceCobNivCaja(pdFecha As Date, pnTipoCambio As Currency, psCtaContCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConect As New DConecta
    oConect.AbreConexion
    sSQL = "Exec stp_sel_GetIndiceCoberturaNivCaja '" & Format(pdFecha, "yyyymmdd") & "'," & pnTipoCambio & ",'" & psCtaContCod & "'"
    Set ObtieneIndiceCobNivCaja = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Public Function GetCtaEfectivoTrans_Age(pdFecha As Date, psMoneda As String, Optional psTpo As String = "") As ADODB.Recordset
Dim sSQL As String
Dim oConect As New DConecta
    oConect.AbreConexion
    sSQL = "Exec stp_sel_GetCtaEfectivoTransito '" & Format(pdFecha, "yyyymmdd") & "','" & psMoneda & "', '" & psTpo & "'"
    Set GetCtaEfectivoTrans_Age = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
'******************END NAGL 20180920*************

