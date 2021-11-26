VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nColocEvalCalProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************
'* Clase que contiene el Proceso de Calificacion de la Cartera
'* LAYG - 01/11/2002
'****************************
Option Explicit
Dim mbTrans As Boolean

Public Function nLLenaTablaCalifProv(ByVal psRutaServer As String, _
        Optional ByVal psFechaData As String = "", Optional ByVal pnTipoCambio As Double = 0) As Recordset

'************************************

Dim lsSQL As String
Dim loReg As DConecta
Dim lrDatos As ADODB.Recordset
Dim lsRFAEstadoVigentes As String
Dim lsRFAEstadoJudicial As String
lsRFAEstadoVigentes = "2060,2061,2062"
lsRFAEstadoJudicial = "2261,2265"

    '*********  Creditos
    lsSQL = "SELECT  C.cCtaCod, C.nSaldoCap, C.nPrdEstado, C.cRefinan, C.nDiasAtraso, " _
         & " C.nMontoApr, C.cLineaCred , C.cCodAnalista, C.dFecVig, " _
         & " cPersCod = (Select Max(cPersCod) From " & psRutaServer & "ProductoPersonaConsol PE " _
         & "             Where PE.cCtaCod=C.cCtaCod And PE.nPrdPersRelac = " & gColRelPersTitular & "), " _
         & "cCalDias = (SELECT t.cCalif cAbrev FROM ColocCalificaTabla t    " _
         & "            WHERE substring(ltrim(str(t.nCalCodTab)),1,3) = Substring(c.cctacod,6,1) + '0' " _
         & "             + Case When C.cRefinan ='R' then '1' else '0' end " _
         & "            AND nRangIni <= C.nDiasAtraso And  nRangFin >= C.nDiasAtraso), " _
         & "nGarPref =( Select Isnull(Sum(case when nMoneda =1 then nHipoteca + nVehicular + + nIndustrial + nAgricola " _
         & "                                   when nMoneda = 2 then (nHipoteca + nVehicular + nIndustrial + nAgricola ) * " & pnTipoCambio & " " _
         & "             end),0) " _
         & "            From " & psRutaServer & "CredGarantias cg " _
         & "            Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "' ), " _
         & "nGarMuyRR = 0, " _
         & "nGarAutoL = ( Select Isnull(sum(case when nMoneda = 1 then nDeposito " _
         & "                                     when nMoneda = 2 then (nDeposito )  * " & pnTipoCambio & " " _
         & "              end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "') " _
         & " FROM  " & psRutaServer & "CreditoConsol C  " _
         & " JOIN " & psRutaServer & "CREDITOSALDOCONSOL CS ON CS.cCtaCod =C.cCtaCod " _
         & " Where C.nPrdEstado in ( " & gColocEstVigNorm & "," & gColocEstVigMor & "," & gColocEstVigVenc & "," _
         & gColocEstRefNorm & "," & gColocEstRefMor & "," & gColocEstRefVenc & "," _
         & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & "," & lsRFAEstadoVigentes & ") " _
         & " And Substring(C.cCtaCod,7,2) <> '21' " _
         & " AND DateDiff(dd,CS.dFecha,'" & Format(psFechaData, "mm/dd/yyyy") & "') = 0 "
        'Recuperaciones
    lsSQL = lsSQL & " Union " _
         & "SELECT  C.cCtaCod, C.nSaldoCap, C.nPrdEstado, C.cRefinan, C.nDiasAtraso, " _
         & " C.nMontoApr, C.cLineaCred , C.cCodAnalista, C.dFecVig, " _
         & " cPersCod = (Select Max(cPersCod) From " & psRutaServer & "ProductoPersonaConsol PE " _
         & "             Where PE.cCtaCod=C.cCtaCod And PE.nPrdPersRelac = " & gColRelPersTitular & "), " _
         & "cCalDias = (SELECT t.cCalif cAbrev FROM ColocCalificaTabla t        " _
         & "            WHERE substring(ltrim(str(t.nCalCodTab)),1,3) = Substring(c.cctacod,6,1) + '0' " _
         & "             + Case When C.cRefinan ='R' then '1' else '0' end " _
         & "            AND nRangIni <= C.nDiasAtraso And  nRangFin >= C.nDiasAtraso), " _
         & "nGarPref =( Select Isnull(Sum(case when nMoneda =1 then nHipoteca + nVehicular + + nIndustrial + nAgricola " _
         & "                                   when nMoneda = 2 then (nHipoteca + nVehicular + nIndustrial + nAgricola ) * " & pnTipoCambio & " " _
         & "             end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "'), " _
         & "nGarMuyRR = 0, " _
         & "nGarAutoL = ( Select Isnull(sum(case when nmoneda = 1 then nCartaFianza + nDeposito " _
         & "                                     when nmoneda = 2 then (nCartaFianza + nDeposito )  * " & pnTipoCambio & " " _
         & "              end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "' ) " _
         & " FROM  " & psRutaServer & "CreditoConsol C  " _
         & " JOIN " & psRutaServer & "CREDITOSALDOCONSOL CS ON CS.cCtaCod =C.cCtaCod " _
         & " Where C.nPrdEstado in ( " & gColocEstRecVigJud & "," & gColocEstRecVigCast & ",2205,2206," & lsRFAEstadoJudicial & " )" _
         & " AND DateDiff(dd,CS.dFecha,'" & Format(psFechaData, "mm/dd/yyyy") & "') = 0 "
    'Carta Fianza
    lsSQL = lsSQL & " Union " _
         & " SELECT  C.cCtaCod, C.nMontoApr as nSaldoCap, nPrdEstado, 'N' as cRefinan, 0 as nDiasAtraso, " _
         & " C.nMontoApr, '0101'+ Substring(C.cCtaCod,9,1) + '1' + Substring(C.cCtaCod,6,3)+'01' as cLineaCred , C.cCodAnalista, C.dFecVig, " _
         & " cPersCod = (Select Max(cPersCod) From " & psRutaServer & "ProductoPersonaConsol PE " _
         & "             Where PE.cCtaCod=C.cCtaCod And PE.nPrdPersRelac = " & gColRelPersTitular & " ) , " _
         & "cCalDias = '0', " _
         & "nGarPref =( Select Isnull(Sum(case when nMoneda =1 then nHipoteca + nVehicular + + nIndustrial + nAgricola " _
         & "                                   when nMoneda = 2 then (nHipoteca + nVehicular + nIndustrial + nAgricola ) * " & pnTipoCambio & " " _
         & "             end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "' ), " _
         & "nGarMuyRR = 0, " _
         & "nGarAutoL = ( Select Isnull(sum(case when nmoneda = 1 then nCartaFianza + nDeposito " _
         & "                                     when nMoneda = 2 then (nCartaFianza + nDeposito )  * " & pnTipoCambio & " " _
         & "              end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "' ) " _
         & " FROM  " & psRutaServer & "CartaFianzaConsol C  " _
         & " Where C.nPrdEstado in ( " & gColocEstVigNorm & "," & gColocEstVigMor & "," & gColocEstVigVenc & ") " _

    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos Is Nothing Then
        MsgBox "No existen datos para Calificacion", vbInformation, "Aviso"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        MsgBox "No existen datos para Calificacion", vbInformation, "Aviso"
        Exit Function
    End If
     
    Set nLLenaTablaCalifProv = lrDatos
    Set lrDatos = Nothing
End Function

'** Asigna calificacion al credito segun los dias de atraso y las tablas de codigos
'** Utiliza ColocCalificaTabla

Public Function nCalifxDiasAtraso(ByVal psCodCta As String, ByVal pnPrdEstado As Integer, _
        ByVal pbGarantPref As Boolean, ByVal pnDiasAtraso As Integer, _
        ByVal psRutaServer As String) As String


Dim rs As ADODB.Recordset
Dim loConecta As DConecta
Dim lsSQL As String
Dim lsCodTab As String
Dim lsCondicProv As String
Dim lsRefinanciado As String
'If Mid(psCodCta, 6, 1) = "1" Or Mid(psCodCta, 6, 1) = "2" Or Mid(psCodCta, 6, 1) = "4" Then
'    lsCondicProv = Mid(psCodCta, 6, 1)
'    lsCondicProv = lsCondicProv & IIf(pbGarantPref = True, "1", "0")  ' 1 Con GarPref / 0 Sin GarPref
'ElseIf Mid(psCodCta, 6, 1) = "3" And Mid(psCodCta, 6, 3) <> "305" Then
'    lsCondicProv = Mid(psCodCta, 6, 1) & "0"
'ElseIf Mid(psCodCta, 6, 3) = "305" Then
'    lsCondicProv = "70"
'End If

'lsCodTab = lsCondicProv
If pnPrdEstado = "2030" Or pnPrdEstado = "2031" Or pnPrdEstado = "2032" _
                        Or pnPrdEstado = "2205" Then
    lsRefinanciado = "1"
Else
    lsRefinanciado = "0"
End If
lsCodTab = Mid(psCodCta, 6, 1) + IIf(pbGarantPref = True, "1", "0") + lsRefinanciado

If pnPrdEstado = gColocEstRecVigJud Or pnPrdEstado = 2205 Then ' Cobranza Judicial
        '// Creditos Judiciales con menos de 120 dias de mora acumulada
        '// Se califican Dudosos - 11/10/2002- Luis Lopez
'    If fgJudicialMora120(lsCodCta) = True Then  ' Si tiene Menos de 120 dias de Mora
'        fgAsinaCalificacionCredito = "3"
'    Else
        nCalifxDiasAtraso = "4"
'    endif
Else
    lsSQL = "Select cCalif From ColocCalificaTabla " _
          & "Where nCalCodTab like '" & lsCodTab & "%'  " _
          & "AND nRangIni <= " & pnDiasAtraso & " AND nRangFin >= " & pnDiasAtraso & " "
    Set loConecta = New DConecta
        loConecta.AbreConexion
        Set rs = loConecta.CargaRecordSet(lsSQL)
    Set loConecta = Nothing
    If (rs.BOF And rs.EOF) Or rs Is Nothing Then
        MsgBox " No encuentro Tabla de Calificaciones", vbInformation, "Aviso"
    Else
        nCalifxDiasAtraso = Trim(rs!cCalif)
    End If
    rs.Close
    Set rs = Nothing
End If
End Function

Public Function nCalifCalculaMontoProvision(ByVal psCodCta As String, ByVal pnSaldoCap As Currency, _
            ByVal psCalificacion As String, ByVal pbJudicial As Boolean, ByVal pbGarantPref As Boolean, _
            ByVal psRutaServer As String, ByVal pnTipoCambioFijo As Double) As Currency
'****
Dim lsSQL As String
Dim lrP As New ADODB.Recordset
Dim loConec As DConecta
Dim lnMontoProvision As Currency

Dim lsCondicProv As String

Dim sGarClase As String
Dim nSaldoCapitalTempo As Double
Dim nPorcPref As Double
Dim nPorcNOPref As Double
Dim cParametro As String

'    Set loConec = New DConecta
'    loConec.AbreConexion
'
'    lsSQL = "Select nConsSisValor from ConstSistema Where nConsSisCod = 75 "
'    Set lrP = loConec.CargaRecordSet(lsSQL)
'    cParametro = Trim(lrP!nConsSisValor)
'    lrP.Close
'
'    If Mid(psCodCta, 6, 1) = "1" Or Mid(psCodCta, 6, 1) = "2" Or Mid(psCodCta, 6, 1) = "4" Then
'        lsCondicProv = Mid(psCodCta, 6, 1)
'        If cParametro <> "3" Then
'            lsCondicProv = lsCondicProv & IIf(pbGarantPref = True, "1", "0") ' 1 Con GarPref / 0 Sin GarPref
'        End If
'    ElseIf Mid(psCodCta, 6, 1) = "3" And Mid(psCodCta, 6, 3) <> "305" Then
'        If cParametro <> "3" Then
'            lsCondicProv = Mid(psCodCta, 6, 1) & "0"
'        Else
'            lsCondicProv = Mid(psCodCta, 6, 1)
'        End If
'    ElseIf Mid(psCodCta, 6, 3) = "305" Then
'        If cParametro <> "3" Then
'            lsCondicProv = "70"
'        Else
'            lsCondicProv = "7"
'        End If
'    End If
'
'    nPorcPref = 0
'    nPorcNOPref = 0
'
'    lsSQL = "Select * From ColocCalificaTabla " _
'      & " Where nCalCodTab like '" & lsCondicProv & "%' " _
'      & " And cCalif = '" & psCalificacion & "' "
'        Set lrP = loConec.CargaRecordSet(lsSQL)
'        If lrP.BOF And lrP.EOF Then
'            MsgBox "No encuentro Porcentaje de Provision en Tablas ", vbInformation, "Aviso"
'        Else
'            If cParametro = "3" Then
'                Do While Not lrP.EOF
'                    If Mid(lrP!nCalCodTab, 2, 1) = "1" Then
'                        nPorcPref = lrP!nProvision
'                    Else
'                        nPorcNOPref = lrP!nProvision
'                    End If
'                    lrP.MoveNext
'                Loop
'            Else
'                lnMontoProvision = Format(pnSaldoCap * (Val(lrP!nProvision) / 100), "#0.0000")
'            End If
'
'        End If
'        lrP.Close
'
'    If cParametro = "3" Then
'
'        lsSQL = "Select G.nGarClase "
'        lsSQL = lsSQL & " from ColocGarantia  CG Inner Join Garantias G ON G.cNumGarant = CG.cNumGarant "
'        lsSQL = lsSQL & " Where cCtaCod =  '" & psCodCta & "' AND (CG.nEstado IS NULL OR CG.nEstado = 1)"
'        lsSQL = lsSQL & " Group by G.nGarClase "
'
'        Set lrP = loConec.CargaRecordSet(lsSQL)
'
'        lnMontoProvision = 0
'        nSaldoCapitalTempo = pnSaldoCap
'        'Si el credito tiene mas de una garantia de diferente tipo de clase (preferida y no preferida)
'        If lrP.RecordCount > 1 Then
'            lrP.Close
'
'            lsSQL = " Select G.nGarClase, G.nGarTpoRealiz, (CASE WHEN CONVERT(char(1),CG.nMoneda) <> SUBSTRING(CG.cCtaCod,9,1) THEN  CASE WHEN CG.nMoneda = 2 THEN CG.nGravado * " & Format(pnTipoCambioFijo, "#0.000") & " ELSE  CG.nGravado / " & Format(pnTipoCambioFijo, "#0.000") & " END ELSE CG.nGravado END) AS nGravado "
'            lsSQL = lsSQL & " from ColocGarantia  CG Inner Join Garantias G ON G.cNumGarant = CG.cNumGarant "
'            lsSQL = lsSQL & " Where cCtaCod =  '" & psCodCta & "' AND (CG.nEstado IS NULL OR CG.nEstado = 1)"
'            lsSQL = lsSQL & " Order By G.nGarClase, G.nGarTpoRealiz DESC "
'            Set lrP = loConec.CargaRecordSet(lsSQL)
'
'            Do While Not lrP.EOF
'                If nSaldoCapitalTempo > lrP!nGravado Then
'                    If lrP!nGarClase = 1 Then
'                        If nPorcPref = 0 Then
'                            MsgBox "Porcentaje de Garantia Preferida para Credito : " & psCodCta & " es Cero o No Existe"
'                        End If
'                        lnMontoProvision = lnMontoProvision + ((nPorcPref / 100) * lrP!nGravado)
'                    Else
'                        lnMontoProvision = lnMontoProvision + ((nPorcNOPref / 100) * lrP!nGravado)
'                        If nPorcNOPref = 0 Then
'                            MsgBox "Porcentaje de Garantia Preferida para Credito : " & psCodCta & " es Cero o No Existe"
'                        End If
'                    End If
'                Else
'                    If lrP!nGarClase = 1 Then
'                        If nPorcPref = 0 Then
'                            MsgBox "Porcentaje de Garantia Preferida para Credito : " & psCodCta & " es Cero o No Existe"
'                        End If
'                        lnMontoProvision = lnMontoProvision + ((nPorcPref / 100) * nSaldoCapitalTempo)
'                    Else
'                        If nPorcNOPref = 0 Then
'                            MsgBox "Porcentaje de Garantia Preferida para Credito : " & psCodCta & " es Cero o No Existe"
'                        End If
'                        lnMontoProvision = lnMontoProvision + ((nPorcNOPref / 100) * nSaldoCapitalTempo)
'                    End If
'                    Exit Do
'                End If
'                nSaldoCapitalTempo = nSaldoCapitalTempo - lrP!nGravado
'                If nSaldoCapitalTempo < 0 Then
'                    nSaldoCapitalTempo = 0
'                End If
'                lrP.MoveNext
'            Loop
'            lrP.Close
'
'            lnMontoProvision = Format(lnMontoProvision, "#0.0000")
'        Else
'            If lrP.RecordCount = 0 Then
'                If Mid(psCodCta, 6, 3) = "305" Then
'                    lnMontoProvision = Format(pnSaldoCap * (nPorcNOPref / 100), "#0.0000")
'                Else
'                    MsgBox "No existe Garantias para el Credito : " & psCodCta
'                End If
'            Else
'                If lrP!nGarClase = 1 Then
'                    If nPorcPref = 0 Then
'                        'MsgBox "Porcentaje de Garantia Preferida para Credito : " & psCodCta & " es Cero o No Existe"
'                    End If
'                    lnMontoProvision = Format(pnSaldoCap * (nPorcPref / 100), "#0.0000")
'                Else
'                    If nPorcNOPref = 0 Then
'                        MsgBox "Porcentaje de Garantia Preferida para Credito : " & psCodCta & " es Cero o No Existe"
'                    End If
'                    lnMontoProvision = Format(pnSaldoCap * (nPorcNOPref / 100), "#0.0000")
'                End If
'            End If
'            lrP.Close
'        End If
'
'        Set lrP = Nothing
'        Set loConec = Nothing
'    End If

nCalifCalculaMontoProvision = lnMontoProvision

End Function

Public Function nCorrigeCalifxDiaAtraso(ByVal psCalif As String, ByVal pnPrdEstado As Integer) As String
Dim lsCalifCorrige As String
lsCalifCorrige = psCalif
    '*** Reglas de Negocio -*** Ica incluye los RFA
    'Asigna la calificacion a refinanciados (Si calif < 2 ==> asigna (2) )
    If pnPrdEstado = gColocEstRefNorm Or pnPrdEstado = 2260 Then  'Vigente Refinanciado
        If Val(psCalif) < 2 Then
            lsCalifCorrige = "2"
        End If
    End If
    'Creditos en Recuperaciones = 4 -*** Ica incluye los RFA
    If pnPrdEstado = gColocEstRecVigJud Or pnPrdEstado = gColocEstRecVigCast Or pnPrdEstado = 2205 Or pnPrdEstado = 2206 Or pnPrdEstado = 2261 Or pnPrdEstado = 2262 Or pnPrdEstado = 2265 Or pnPrdEstado = 2266 Then       'Vigente Refinanciado
        lsCalifCorrige = "4"
    End If
nCorrigeCalifxDiaAtraso = lsCalifCorrige
End Function


Public Function ObtieneTablaCalif(ByVal psServer As String) As ADODB.Recordset
Dim sql As String
Dim Co As DConecta
Dim rs As ADODB.Recordset
Set Co = New DConecta
sql = "Select nCalCodTab, cCalDes, nRangIni, nRangfin, cCalif, nProvision, cRefinan " & _
      "From ColocCalificaTabla Order by nCalCodTab"
Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
Co.CierraConexion
Set ObtieneTablaCalif = rs
End Function


Public Sub InsertaColocCalifTabla(ByVal psServer As String, ByVal psCalDes As String, _
 ByVal pnRangIni As Integer, ByVal pnRangFin As Integer, ByVal psCalif As String, ByVal pnProvision As Double)
Dim Co As DConecta
Dim Sql1 As String
Dim Sql2 As String
Dim rs As New ADODB.Recordset
Dim nNro As Integer
Set Co = New DConecta
Sql1 = "Select Isnull(max(nCalCodTab),0) Nro From ColocCalificatabla "
Co.AbreConexion
Set rs = Co.CargaRecordSet(Sql1)
nNro = nNro + 1
Sql2 = " Insert ColocCalificaTabla (nCalCodTab, cCalDes, nRangIni, nRangfin, cCalif, nProvision) " & _
       "  values (" & nNro & ",'" & psCalDes & "'," & pnRangIni & "," & _
       " " & pnRangFin & ",'" & psCalif & "'," & pnProvision & ")"
Co.Ejecutar (Sql2)
Co.CierraConexion
Set rs = Nothing
Set Co = Nothing
End Sub


Public Sub UpdateColocCalifTabla(ByVal psServer As String, ByVal psCalDes As String, _
 ByVal pnRangIni As Integer, ByVal pnRangFin As Integer, ByVal psCalif As String, _
 ByVal pnProvision As Double, ByVal pnCod As Integer)
Dim Co As DConecta
Set Co = New DConecta
Dim sql As String
sql = " Update ColocCalificaTabla  " & _
      " Set  cCalDes='" & psCalDes & "', nRangIni=" & pnRangIni & ", " & _
      " nRangfin=" & pnRangFin & ", cCalif='" & psCalif & "', nProvision=" & pnProvision & "  " & _
      " Where  nCalCodTab = " & pnCod
Co.AbreConexion
Co.Ejecutar (sql)
Co.CierraConexion
Set Co = Nothing
End Sub


