VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCtasaldo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A82E291008C"
Option Base 0
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sCentralCom As String
Public Function PermiteMntSdoInicial() As Boolean
   On Error GoTo GetFechaSdoInicialErr
   psSql = "SELECT TOP 2 dCtaSaldoFecha FROM CtaSaldo GROUP BY dCtaSaldoFecha "
   Set prs = dbConec.CargaRecordSet(psSql)
   If Not prs.EOF Then
      If prs.RecordCount > 1 Then
         PermiteMntSdoInicial = False
      Else
         PermiteMntSdoInicial = True
      End If
   Else
      PermiteMntSdoInicial = True
   End If
   prs.Close: Set prs = Nothing
   Exit Function
GetFechaSdoInicialErr:
   Call RaiseError(MyUnhandledError, "NCtaSaldo:GetFechaSdoInicial Method")
End Function

Public Function GetFechaSdoInicial() As String
   On Error GoTo GetFechaSdoInicialErr
   psSql = "SELECT MIN(dCtaSaldoFecha) dCtaSaldoFecha FROM CtaSaldo"
   Set prs = dbConec.CargaRecordSet(psSql)
   If Not prs.EOF Then
      If IsNull(prs!dCtaSaldoFecha) Then
         GetFechaSdoInicial = "  /  /    "
      Else
         GetFechaSdoInicial = Format(prs!dCtaSaldoFecha, gsFormatoFechaView)
      End If
   Else
      GetFechaSdoInicial = "  /  /    "
   End If
   prs.Close: Set prs = Nothing
   Exit Function
GetFechaSdoInicialErr:
   Call RaiseError(MyUnhandledError, "NCtaSaldo:GetFechaSdoInicial Method")
End Function

Public Function ExisteCuentaSaldo(ByVal psCtaCod As String, ByVal psFecha As String) As Boolean
   On Error GoTo ExisteCuentaErr
   Dim clsSdo As New DCtaSaldo
   ExisteCuentaSaldo = False
   Set prs = clsSdo.CargaCtaSaldo(psCtaCod, psFecha)
     If Not (prs.BOF And prs.EOF) Then
         ExisteCuentaSaldo = True
     End If
   prs.Close: Set prs = Nothing
   Set clsSdo = Nothing
   Exit Function
ExisteCuentaErr:
   Call RaiseError(MyUnhandledError, "NCtaSaldo:ExisteCuenta Method")
End Function

Public Sub CierreContableMensual(psFecIni As String, psFecFin As String, psCodUser As String, psFechaAct As String)
Dim sFec As String
Dim sAge As String
Dim sTpo As String
Dim sClave  As String
Dim N       As Currency
Dim sMovNro As String
Dim oCon    As DConecta
Dim oMov As DMov
On Error GoTo CierreContableMensualErr
   Dim oGen As New NConstSistemas
   Set oMov = New DMov
   sMovNro = oMov.GeneraMovNro(CDate(Format(psFechaAct, gsFormatoFechaView)), , psCodUser)
   Set oGen = Nothing

'Primero Creamos MovCont para todos los Mov Contables
Set oCon = New DConecta
oCon.AbreConexion
psSql = "INSERT MovCont " _
      & "SELECT DISTINCT m.nMovNro, 0, 0, '' " _
      & "From Mov M LEFT JOIN MovCont M1 ON M1.nMovNro = M.nMovNro " _
      & "        JOIN MovCta mc ON mc.nMovNro = m.nMovNro " _
      & "Where M.nMovEstado = " & gMovEstContabMovContable & " and not M.nMovFlag IN (" & gMovFlagEliminado & "," & gMovFlagModificado & ") and " _
      & "      substring(M.cMovNro,1,8) between '" & psFecIni & "' and '" & psFecFin & "' " _
      & "      and mc.cCtaContCod <> '' and m1.nMovNro Is NULL "
oCon.Ejecutar psSql
oCon.CierraConexion
Set oCon = Nothing
'   psSql = "Select M.cMovNro, M.nMovNro, ISNULL(O.cOpeGruCod,'') as cTpo " _
'        & "From Mov M Join MovCta MC On M.cMovNro = MC.cMovNro " _
'        & "           Join " & sCentralCom & "CtaCont C On MC.cCtaContCod = C.cCtaContCod " _
'        & "           Join " & sCentralCom & "OpeTpo  O On M.cOpeCod = O.cOpeCod " _
'        & "Where M.nMovEstado = " & gMovEstContabMovContable & " and M.nMovFlag <> " & gMovFlagEliminado & " and convert(datetime,LEFT(M.cMovNro,8),101) between " _
'        & "      '" & psFecIni & "' And '" & psFecFin & "' AND mc.cCtaContCod <> '' " _
'        & "GROUP BY M.cMovNro, ISNULL(O.cOpeGruCod,'') " _
'        & "Order by LEFT(M.cMovNro,8), substring(M.cMovNro,18,2), ISNULL(O.cOpeGruCod,'') "
   
   psSql = "spCierreContMensual '" & psFecIni & "', '" & psFecFin & "'"
   Set prs = dbConec.CargaRecordSet(psSql)
   oGen.ActualizaConstSistemas gConstSistCierreMensualCont, sMovNro, Right(psFecFin, 2) & "/" & Mid(psFecFin, 5, 2) & "/" & Left(psFecFin, 4)
Exit Sub
CierreContableMensualErr:
   Call RaiseError(MyUnhandledError, "NCtaSaldo:CierreContableMensual Method")
End Sub

Public Function GetCtaSaldo(psCtaContCod As String, psFecha As String, Optional pbMN As Boolean = True) As Currency
   On Error GoTo GetCtaSaldoErr
   Dim lsCond As String
   If Right(psCtaContCod, 1) = "%" Then
        lsCond = " LIKE "
   Else
        lsCond = " = "
   End If
   
   If lsCond <> " = " Then
        If Mid(psCtaContCod, 3, 1) = "0" Then
            psSql = "Select dbo.GetSaldoCtaAcumulado('" & psFecha & "','" & Left(psCtaContCod, 2) & "_" & Mid(psCtaContCod, 4) & "'," & IIf(pbMN, "1", "2") & ") as Saldo"
            Set prs = dbConec.CargaRecordSet(psSql)
            If Not prs.EOF Then
               If pbMN Then '
                  GetCtaSaldo = prs!Saldo
               Else
                  GetCtaSaldo = prs!Saldo
               End If
            Else
               GetCtaSaldo = 0
            End If
            
        Else
            psSql = "SELECT ISNULL(SUM(nCtaSaldoImporte),0) nCtaSaldoImporte, ISNULL(SUM(nCtaSaldoImporteME),0) nCtaSaldoImporteME " _
                  & "FROM CtaSaldo cs WHERE cCtaContCod " & lsCond & "'" & psCtaContCod & "' AND dCtaSaldoFecha = " _
                  & "      (SELECT MAX(dCtaSaldoFecha) FROM CtaSaldo cs1 WHERE cs1.cCtaContCod = cs.cCtaContCod and dCtaSaldoFecha <= '" & psFecha & "')"
        
             Set prs = dbConec.CargaRecordSet(psSql)
             If Not prs.EOF Then
                If pbMN Then
                   GetCtaSaldo = prs!nCtaSaldoImporte
                Else
                   GetCtaSaldo = prs!nCtasaldoImporteME
                End If
             Else
                GetCtaSaldo = 0
             End If
        End If
    Else
       If pbMN Then
          psSql = "SELECT dbo.getsaldocta('" & psFecha & "','" & psCtaContCod & "',1) "
       Else
          psSql = "SELECT dbo.getsaldocta('" & psFecha & "','" & psCtaContCod & "',2) "
       End If
    
       Set prs = dbConec.CargaRecordSet(psSql)
       GetCtaSaldo = prs.Fields(0)
       
    End If
   Exit Function
GetCtaSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:GetCtaSaldo Method")
End Function

Public Function GetOpeCtaSaldo(psOpeCod As String, psFecha As String, Optional pbMN As Boolean = True, Optional psOpeCtaOrden As String = "", Optional psFiltro As String = "") As Currency
Dim sFiltro As String
   On Error GoTo GetOpeCtaSaldoErr
   If psOpeCtaOrden <> "" Then
      sFiltro = " and cOpeCtaOrden ='" & psOpeCtaOrden & "'  "
   End If
   psSql = "SELECT    dCtaSaldoFecha , SUM(nCtaSaldoImporte) as SaldoMN, SUM(nCtaSaldoImporteME) as SaldoME   " _
         & " FROM     Opecta  OC " _
         & "          JOIN CTASALDO CS ON CS.cCtaContCod LIKE LEFT(OC.CCTACONTCOD,2)+'" & IIf(pbMN, "[136]", "2") & "'+SubString(OC.cCtaContCod,4,22) + '%' " _
         & " WHERE    cOpecod = '" & psOpeCod & "' AND nCtaSaldoImporte<>0 " _
         & "          AND dCtaSaldoFecha =(   SELECT MAX(dCtaSaldoFecha) " _
         & "                                  FROM CTASALDO CS1 " _
         & "                                  WHERE CS1.cCtaContCod =CS.cCtaContCod AND dCtaSaldoFecha <= '" & psFecha & "') " _
         & sFiltro & psFiltro _
         & " GROUP BY dCtaSaldoFecha "
            
   Set prs = dbConec.CargaRecordSet(psSql)
   If Not prs.EOF Then
    Do While Not prs.EOF
       GetOpeCtaSaldo = GetOpeCtaSaldo + IIf(pbMN, prs!SaldoMN, prs!SaldoME)
        prs.MoveNext
    Loop
   Else
      GetOpeCtaSaldo = 0
   End If
   Exit Function
GetOpeCtaSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:GetOpeCtaSaldo Method")
End Function

Public Function GetCtaSaldoRango(psCtaContCod As String, pdFecIni As Date, pdFecFin As Date) As ADODB.Recordset
   On Error GoTo GetCtaSaldoErr
Dim lsCondCta As String

'Probado por Pepe En Lima, Trujillo, Chimbote e Ica

If InStr(psCtaContCod, ",") > 0 Then
    lsCondCta = " IN (" & psCtaContCod & ") "
Else
    lsCondCta = " like '" & psCtaContCod & "' "
End If

psSql = "SELECT Fec.dFecha, Fec.dFecha, ISNULL(SUM(nCtaSaldoImporte),0) nCtaSaldoImporte, ISNULL(SUM(nCtaSaldoImporteME),0) nCtaSaldoImporteME "
psSql = psSql & " FROM FechaTmp('" & Format(pdFecIni, gsFormatoFecha) & "') Fec, CTASALDO CS1 "
psSql = psSql & " WHERE cs1.cCtaContCod " & lsCondCta & " "
psSql = psSql & "       and CS1.dCtaSaldoFecha = ( "
psSql = psSql & "                                   SELECT  MAX(dCtaSaldoFecha) "
psSql = psSql & "                                   FROM   CtaSaldo CS "
psSql = psSql & "                                   Where cs1.cCtaContCod = CS.cCtaContCod "
psSql = psSql & "                                         AND CS.dCtaSaldoFecha <= Fec.dFecha ) "
psSql = psSql & " and Convert(varchar(8), Fec.dFecha, 112) >='" & Format(pdFecIni, "YYYYmmdd") & "' "
psSql = psSql & " and Convert(varchar(8), Fec.dFecha, 112) <='" & Format(pdFecFin, "YYYYmmdd") & "' "
    
psSql = psSql & " GROUP BY Fec.dFecha ORDER BY Fec.dFecha "

Set GetCtaSaldoRango = dbConec.CargaRecordSet(psSql)
 
   Exit Function
GetCtaSaldoErr:
   Call RaiseError(MyUnhandledError, "NCtaSaldo:GetCtaSaldoRango Method")
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As New ClasIni
   sCentralCom = oIni.BaseComunes
Set oIni = Nothing

Set dbConec = New DConecta
dbConec.AbreConexion

End Sub

Private Sub Class_Terminate()
dbConec.CierraConexion
Set dbConec = Nothing
End Sub

Public Function GetCtaEfectivoSaldo(psCtaContCod As String, psFecha As String, Optional pbMN As Boolean = True, Optional psEfectivoCod As String = "") As Currency
Dim lsEfectivoFiltro As String
On Error GoTo GetCtaEfectivoSaldoErr
If psEfectivoCod <> "" Then
   lsEfectivoFiltro = " and cEfectivoCod = '" & psEfectivoCod & "' "
End If
   psSql = "SELECT ISNULL(SUM(nEfectivoSaldoImporte),0) nCtaSaldoImporte, ISNULL(SUM(nEfectivoSaldoImporteME),0) nCtaSaldoImporteME " _
         & "FROM CtaObjEfectivoSaldo cs WHERE cCtaContCod LIKE '" & psCtaContCod & "' AND dEfectivoSaldoFecha = " _
         & "      (SELECT MAX(dEfectivoSaldoFecha) FROM CtaObjEfectivoSaldo cs1 WHERE cs1.cCtaContCod = cs.cCtaContCod and cs1.cEfectivoCod=cs.cEfectivoCod and dEfectivoSaldoFecha <= '" & psFecha & "') " _
         & lsEfectivoFiltro
   Set prs = dbConec.CargaRecordSet(psSql)
   If Not prs.EOF Then
      If pbMN Then
         GetCtaEfectivoSaldo = prs!nCtaSaldoImporte
      Else
         GetCtaEfectivoSaldo = prs!nCtasaldoImporteME
      End If
   Else
      GetCtaEfectivoSaldo = 0
   End If
   Exit Function
GetCtaEfectivoSaldoErr:
   Call RaiseError(MyUnhandledError, "NCtaSaldo:GetCtaEfectivoSaldo Method")
End Function

'BRGO 20110314 ********************************************************************
'Public Function GetGeneraDatosRep25(ByVal psFecha As String, ByRef rsDism As ADODB.Recordset, ByRef rsRecReexME As ADODB.Recordset)
Public Function GetGeneraDatosRep25(ByVal lsMes As String, ByVal lsAnio As String, ByRef rsDism As ADODB.Recordset, ByRef rsRecReexME As ADODB.Recordset) '** PASI 20130914
'Dim lsAnio As String
'Dim lsMes As String

On Error GoTo GetGeneraDatosRep25Err

'lsAnio = Left(psFecha, 4)
'lsMes = Mid(psFecha, 5, 2)
'   psSql = " Select TipoRep, SaldoIni = SUM(SaldoIni), SaldoFin = SUM(SaldoFin), "
'   psSql = psSql & "  Corpor = Sum(Case TipoCred when '10' Then Debe else 0 end), "
'   psSql = psSql & "  GraEmp = Sum(Case TipoCred when '11' Then Debe else 0 end), "
'   psSql = psSql & "  MedEmp = Sum(Case TipoCred when '12' Then Debe else 0 end), "
'   psSql = psSql & "  PeqEmp = Sum(Case TipoCred when '13' Then Debe else 0 end), "
'   psSql = psSql & "  MicEmp = Sum(Case TipoCred when '02' Then Debe else 0 end), "
'   psSql = psSql & "  Consum = Sum(Case TipoCred when '03' Then Debe else 0 end), "
'   psSql = psSql & "  Dismin = Sum(Haber)"
'   psSql = psSql & "  From ("
'   psSql = psSql & "  Select TipoRep, RIGHT(A.cCtaContCod,2) TipoCred, "
'   psSql = psSql & "      SaldoIni = sum(Case When A.cBalanceMes = " & lsMes & "-2 Then A.nSaldoIniImporte Else 0 End), "
'   psSql = psSql & "      SaldoFin = sum(Case When A.cBalanceMes = " & lsMes & " Then A.nSaldoFinImporte Else 0 End), "
'   psSql = psSql & "      Debe = Sum(A.nDebe), Haber = Sum(A.nHaber) "
'   psSql = psSql & "  from ( "
'   psSql = psSql & "          Select '1' TipoRep, cBalanceMes, nSaldoIniImporte, cCtaContCod, nDebe, nHaber, nSaldoFinImporte "
'   psSql = psSql & "          From BalanceEstad "
'   psSql = psSql & "              Where cBalanceAnio = " & lsAnio & " And cBalanceMes in (" & lsMes & "-2," & lsMes & "-1," & lsMes & ") and cBalanceCate = 1 "
'   psSql = psSql & "                  And cBalanceTipo = 0 and LEN(cCtaContCod) = 8 and cCtaContCod like '810302%' "
'   psSql = psSql & "          Union "
'   psSql = psSql & "          select '1' TipoRep, cBalanceMes, nSaldoIniImporte, cCtaContCod, nDebe, nHaber, nSaldoFinImporte "
'   psSql = psSql & "          From BalanceEstad "
'   psSql = psSql & "          where cBalanceAnio = " & lsAnio & " and cBalanceMes in (" & lsMes & "-2," & lsMes & "-1," & lsMes & ") and cBalanceCate = 1 "
'   psSql = psSql & "              And cBalanceTipo = 0 and LEN(cCtaContCod) = 10 and cCtaContCod like '81092501%' "
'   psSql = psSql & "          Union"
'   psSql = psSql & "          select '2' TipoRep, cBalanceMes, nSaldoIniImporte, cCtaContCod, nDebe, nHaber, nSaldoFinImporte "
'   psSql = psSql & "              From BalanceEstad "
'   psSql = psSql & "              where cBalanceAnio = " & lsAnio & " and cBalanceMes in (" & lsMes & "-2," & lsMes & "-1," & lsMes & ") and cBalanceCate = 1 "
'   psSql = psSql & "                  And cBalanceTipo = 0 and LEN(cCtaContCod) = 10 and cCtaContCod like '81092201%' "
'   psSql = psSql & "           ) A group by TipoRep, RIGHT(A.cCtaContCod,2) "
'   psSql = psSql & "          ) B group by TipoRep "

   psSql = "exec stp_sel_ObtenerDatosRep2501 '" & lsMes & "', '" & lsAnio & "'" '** PASI 20130914
   Set rsDism = dbConec.CargaRecordSet(psSql)
   
'   psSql = "  SELECT Recuperaciones = SUM(CASE WHEN Tipo=1 THEN Valor ELSE 0 END), "
'   psSql = psSql & " Reexpresion    = SUM(CASE WHEN Tipo=2 THEN Valor ELSE 0 END), "
'   psSql = psSql & " TotalDisminuc  = SUM(CASE WHEN Tipo=3 THEN Valor ELSE 0 END) "
'   psSql = psSql & " FROM ( "
'   psSql = psSql & " SELECT Tipo = 1, Valor = SUM(nHaber * CASE WHEN SUBSTRING(cCtaCnt,3,1) = 1 THEN 1 ELSE dbo.GetTipoCambioFinMes(1,'" & psFecha & "') END)"
'   psSql = psSql & " FROM AsientoDN A"
'   psSql = psSql & "    INNER JOIN OpeTpo O ON A.copecod=O.copecod"
'   psSql = psSql & "    INNER JOIN Mov M ON M.nMovNro= A.nMovNro"
'   psSql = psSql & " WHERE  MONTH(dfecha) IN (" & lsMes & "-2," & lsMes & "-1," & lsMes & ")"
'   psSql = psSql & "    AND (cCtaCnt like '81[12]302%' OR cCtaCnt like '81[12]925%' OR cCtaCnt like '81[12]92201%')"
'   psSql = psSql & "    AND nMovFlaG = 0 and O.cOpeCod = 130401"
'   psSql = psSql & " UNION "
'   psSql = psSql & " SELECT Tipo = 2, Valor = SUM(Reexp)  FROM ( "
'   psSql = psSql & " SELECT Left(A.cCtaContCod,16) cCtaContCod, C.cCtaContDesc, "
'   psSql = psSql & "     Reexp = ISNULL(SUM(CASE WHEN A.nMovImporte > 0 THEN A.nMovImporte END),0) - ISNULL(SUM(CASE WHEN A.nMovImporte < 0 THEN A.nMovImporte * -1 END),0) "
'   psSql = psSql & " FROM Mov M "
'   psSql = psSql & "     JOIN MovCta A ON a.nMovNro = M.nMovNro "
'   psSql = psSql & "     LEFT JOIN MovME Me ON me.nMovNro = A.nMovNro and Me.nMovItem = A.nMovItem "
'   psSql = psSql & "     JOIN CtaCont C ON C.cCtaContCod = LEFT(A.cCtaContCod,16) "
'   psSql = psSql & "     LEFT JOIN MovDoc Doc ON Doc.nMovNro = M.nMovNro "
'   psSql = psSql & "     LEFT JOIN Documento Docum ON Docum.nDocTpo = Doc.nDocTpo "
'   psSql = psSql & " WHERE  M.nMovEstado = '10' AND NOT M.nMovFlag IN ('1','5','2') and M.cOpecod ='701201' "
'   psSql = psSql & "     AND (a.cCtaContCod Like '812302%' OR A.cCtaContCod LIKE '81292501%') "
'   psSql = psSql & "     AND SUBSTRING(M.cMovNro,1,4) = '" & lsAnio & "' AND SUBSTRING(M.cMovNro,5,2) IN (" & lsMes & "-2," & lsMes & "-1," & lsMes & ") "
'   psSql = psSql & " GROUP BY LEFT(a.cCtaContCod,16),A.cCtaContcod, C.cCtaContDesc, M.cMovNro, M.nMovNro "
'   psSql = psSql & " )A "
'   psSql = psSql & " UNION "
'   psSql = psSql & " SELECT Tipo = 3, Valor = SUM(ISNULL( CASE WHEN mc.nMovImporte < 0 THEN mc.nMovImporte * -1 "
'   psSql = psSql & "     WHEN mc.nMovImporte > 0 THEN 0 END, 0)) "
'   psSql = psSql & " FROM   MovCta MC JOIN Mov M ON M.nMovNro = mc.nMovNro"
'   psSql = psSql & "    JOIN CtaCont c ON c.cCtaContCod = mc.cCtaContCod"
'   psSql = psSql & "    JOIN CtaContClase cls ON mc.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%'"
'   psSql = psSql & "    JOIN OpeTpo O on O.cOpeCod = M.cOpeCod"
'   psSql = psSql & " WHERE M.nMovEstado = 10 and not M.nMovFlag in (1,5,2)"
'   psSql = psSql & "    And SubString(m.cMovNro,1,4) ='" & lsAnio & "' AND SUBSTRING(M.cMovNro,5,2) IN (" & lsMes & "-2," & lsMes & "-1," & lsMes & ")"
'   psSql = psSql & "  And MC.cCtaContCod IS NOT NULL and MC.cCtaContCod <> '' and not M.cOpeCod LIKE '70185%'"
'   psSql = psSql & "  And (MC.cCtaContCod like '81_302%' or MC.cCtaContCod like '81_92501%')"
'   psSql = psSql & "  GROUP BY M.cOpeCod,O.cOpeDesc"
'   psSql = psSql & " )B "
   psSql = "exec stp_sel_ObtenerDatosRep2502 '" & lsMes & "', '" & lsAnio & "'" '**PASI 20130914
   Set rsRecReexME = dbConec.CargaRecordSet(psSql)
    
   Exit Function
GetGeneraDatosRep25Err:
    Call RaiseError(MyUnhandledError, "NCtaSaldo:GetCtaEfectivoSaldo Method")
End Function
'**************************************************************************************

'** BRGO 20110505 ***********
Public Function GetGeneraDatosAnexo4B(ByVal psFecha As Date) As ADODB.Recordset
    On Error GoTo GetGeneraDatosAnx4BErr
        Dim sSql As String
        sSql = "stp_sel_Anexo4B_BienesAdjudicadosRecup '" & Format(psFecha, "MM/dd/yyyy") & "'"
        Set GetGeneraDatosAnexo4B = dbConec.CargaRecordSet(sSql)
        Exit Function
GetGeneraDatosAnx4BErr:
        Call RaiseError(MyUnhandledError, "NCtaSaldo:GetCtaEfectivoSaldo Method")
End Function

Public Function GetGeneraDatosAnexo4A(ByVal psFecha As Date) As ADODB.Recordset
    On Error GoTo GetGeneraDatosAnx4AErr
        Dim sSql As String
        sSql = "stp_sel_Anexo4A_BienesAdjudicados '" & Format(psFecha, "MM/dd/yyyy") & "'"
        Set GetGeneraDatosAnexo4A = dbConec.CargaRecordSet(sSql)
        Exit Function
GetGeneraDatosAnx4AErr:
        Call RaiseError(MyUnhandledError, "NCtaSaldo:GetCtaEfectivoSaldo Method")
End Function
'*** End BRGO ****************

'ALPA 20130504****************
Public Function GetCtaSaldoBalanceME(ByVal psCtaContCod As String, ByRef pnSaldoInicial As Currency, ByVal psFechaDel As String, ByVal psFechaAl As String, ByVal psMoneda As String, Optional pnImporte As Currency = 0, Optional psFiltro As String = "=", Optional pbME As Boolean = False, Optional pbVerDoc As Boolean = True, Optional pbVerPersona As Boolean = False, Optional psTabla As String = "CtaCont", Optional psTipoCta As String = "D") As Currency
    Dim lnSaldo As Currency
    Dim lrs As New ADODB.Recordset
    Dim rsCta As New ADODB.Recordset
    Dim oSdo  As New DCtaSaldo
    Dim oTCambio As New nTipoCambio
    Dim lsMes As String
    Dim lsAnio As String
    Dim ldFechaDelNuevoMA As Date
    Dim ldFechaDelNuevo As Date
    Dim oCont As New NContAsientos
    Dim lnSaldoInicialME As Currency
    Dim nVarTCambio As Integer
    Dim nTipoCambio As Currency
    Dim lnMovNro As String 'MIOL20130802****
On Error GoTo GetCtaSaldoBalanceMEErr
    
    GetCtaSaldoBalanceME = 0
    psFechaDel = Format(psFechaDel, "YYYY/MM/DD") 'MIOL20130802 - "YYYY/DD/MM"
    ldFechaDelNuevoMA = DateAdd("d", -Day(psFechaDel), psFechaDel)
    ldFechaDelNuevo = ldFechaDelNuevoMA
    lsAnio = CStr(Year(ldFechaDelNuevoMA))
    lsMes = CStr(Month(ldFechaDelNuevoMA))
    lsMes = IIf(Len(lsMes) = 1, "0" & lsMes, lsMes)
    'TCCompra
    'TCVenta
    
    If Replace(psCtaContCod, "%", "") = "2114020101" Then
        nVarTCambio = TCFijoMes
    ElseIf (Mid(psCtaContCod, 1, 1) = "4" Or (Mid(psCtaContCod, 1, 2) = "63" Or Mid(psCtaContCod, 1, 2) = "65")) Then
          nVarTCambio = TCVenta
    ElseIf (Mid(psCtaContCod, 1, 1) = "5" Or (Mid(psCtaContCod, 1, 2) = "62") Or (Mid(psCtaContCod, 1, 2) = "64")) Then
          nVarTCambio = TCCompra
    ElseIf Replace(psCtaContCod, "%", "") = "25141912" Then
        nVarTCambio = TCCompra
    Else
        nVarTCambio = TCFijoMes
    End If
    
    nTipoCambio = oTCambio.EmiteTipoCambio(DateAdd("d", 2, ldFechaDelNuevoMA), nVarTCambio)
    
    Set lrs = oSdo.GetSaldoFinalBalance(Replace(psCtaContCod, "%", ""), lsAnio, lsMes, psMoneda)
    If Not (lrs.BOF Or lrs.EOF) Then
        lnSaldoInicialME = lrs!nSaldoFinImporte / nTipoCambio
    End If
    Set oSdo = Nothing
    
    ldFechaDelNuevoMA = DateAdd("d", 1, ldFechaDelNuevo)
    
       ' nDebeD = 0: nHaberD = 0
    If DateDiff("d", ldFechaDelNuevoMA, psFechaDel) <> 0 Then
    Set rsCta = oCont.GetMayorCuenta(psCtaContCod, Format(ldFechaDelNuevoMA, gsFormatoMovFecha), Format(DateAdd("d", -1, psFechaDel), gsFormatoMovFecha), , psFiltro, pbME, , , psTabla)
    If Not (rsCta.BOF Or rsCta.EOF) Then
        Do While Not rsCta.EOF
            lnMovNro = rsCta!nMovNro 'MIOL20130802 *****
            If psTipoCta = "D" Then
                lnSaldoInicialME = lnSaldoInicialME + rsCta!nDebeME - rsCta!nHaberME
            Else
                lnSaldoInicialME = lnSaldoInicialME + rsCta!nHaberME - rsCta!nDebeME
            End If
            rsCta.MoveNext
            'MIOL20130802 *************
            If rsCta.EOF Then
                Exit Do
            End If
            Do While lnMovNro = rsCta!nMovNro
                rsCta.MoveNext
                 If rsCta.EOF Then
                    Exit Do
                 End If
            Loop
            'END MIOL *****************
        Loop
    End If
    End If
    GetCtaSaldoBalanceME = lnSaldoInicialME
Exit Function
GetCtaSaldoBalanceMEErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:GetCtaSaldo Method")

End Function
'*****************************
