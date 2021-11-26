VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nCajero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsBaseComunes As String
Dim vsBasePesonas As String

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oIni = New ClasIni
vsBaseComunes = oIni.BaseComunes
vsBasePesonas = oIni.BasePersonas
Set oIni = Nothing

End Sub
Public Function GrabaHabilitaAgencia(ByVal psFormatoFecha As String, _
                                ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                ByVal pnImporte As Currency, _
                                ByVal psCtaContHaber As String, ByVal psAreaCod As String, ByVal psAgeCod As String, _
                                ByVal psCtaContDebe As String, ByVal psPersCodUser As String, _
                                ByVal psUser As String) As Integer
Dim oMov As DMov
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New DMov

On Error GoTo GrabaHabilitaAgenciaErr
oMov.BeginTrans
lbTrans = True
GrabaHabilitaAgencia = 1
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebe, pnImporte
lnMovOrden = lnMovOrden + 1
oMov.InsertaMovGasto lnMovNro, psPersCodUser, ""

lnMovItem = lnMovItem + 1: lnMovOrden = 0
oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber, pnImporte * -1
lnMovOrden = lnMovOrden + 1
oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
oMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, psAgeCod, psAreaCod

oMov.InsertaMovHabDev lnMovNro, psUser, pnImporte
oMov.CommitTrans
lbTrans = False
GrabaHabilitaAgencia = 0
Set oMov = Nothing
Exit Function
GrabaHabilitaAgenciaErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetMovHabBovAgeCajero(ByVal psOpeCod As String, ByVal psAreaCod As String, _
                                      ByVal psAgeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                      Optional ByVal psCodUser As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim lsFiltro As String

Set oCon = New DConecta
Set rs = New Recordset
If oCon.AbreConexion = False Then Exit Function
lsFiltro = ""
If psCodUser <> "" Then
    lsFiltro = " AND D.CUSER ='" & psCodUser & "'  "
End If

sql = "SELECT CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2)  as Fecha, " _
    & "         O.cOpeDesc ,D.CUSER,  D.CPERSNOMBRE, D.IMPORTE, m.CMOVDESC, m.nMovNro, D.CPERSCOD " _
    & " FROM    MOV M " _
    & "         JOIN ( SELECT   M.CMOVNRO, M.nMOVNRO,  ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)) AS IMPORTE, " _
    & "                         MOP.CPERSCOD , RH.CUSER, P.CPERSNOMBRE, MO.COBJETOCOD " _
    & "                 FROM    MOV M " _
    & "                         JOIN MOVCTA MC  ON MC.nMOVNRO =M.nMOVNRO " _
    & "                         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "                         JOIN MOVOBJ MO ON MO.nMOVNRO = MC.nMOVNRO AND MO.nMOVITEM = MC.nMOVITEM AND MO.COBJETOCOD ='" & Format(ObjPersona, "00") & "' " _
    & "                         JOIN MOVGasto MOP ON MOP.nMOVNRO = MO.nMOVNRO " _
    & "                         JOIN RRHH RH ON RH.CPERSCOD = MOP.CPERSCOD " _
    & "                         JOIN " & vsBasePesonas & "PERSONA P ON P.CPERSCOD = MOP.CPERSCOD " _
    & "                 WHERE   M.COPECOD ='" & psOpeCod & "' AND MC.CCTACONTCOD>0 ) AS D ON D.nMOVNRO = M.nMOVNRO " _
    & "         JOIN (  SELECT  M.CMOVNRO,M.nMOVNRO,  MOA.CAREACOD, MOA.CAGECOD, MO.COBJETOCOD " _
    & "                 FROM    MOV M " _
    & "                         JOIN MOVCTA MC  ON MC.nMOVNRO =M.nMOVNRO " _
    & "                         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "                         JOIN MOVOBJ MO ON MO.nMOVNRO = MC.nMOVNRO AND MO.nMOVITEM = MC.nMOVITEM AND MO.COBJETOCOD ='" & ObjCMACAgenciaArea & "' " _
    & "                         JOIN MOVOBJAREAAGENCIA MOA ON MOA.nMOVNRO = MO.nMOVNRO AND MOA.nMOVITEM = MO.nMOVITEM AND MOA.nMOVOBJORDEN = MO.nMOVOBJORDEN " _
    & "                 WHERE   M.COPECOD ='" & psOpeCod & "') AS H ON H.nMOVNRO = M.nMOVNRO " _
    & "         JOIN OPETPO O ON O.COPECOD = M.COPECOD " _
    & " WHERE   M.COPECOD ='" & psOpeCod & "' AND M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagEliminado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
    & "         AND H.CAREACOD ='" & psAreaCod & "' AND H.CAGECOD='" & psAgeCod & "' " _
    & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' " & lsFiltro
sql = sql + "   AND NOT EXISTS (    SELECT  M1.CMOVNRO " _
    & "                             FROM    MOV M1 JOIN MOVREF MR ON MR.nMOVNRO = M1.nMOVNRO " _
    & "                             WHERE   M1.nMovFlag IN (" & gMovFlagVigente & ") AND MR.nMOVNROREF = M.nMOVNRO ) " _
    & " ORDER BY M.CMOVNRO   "
 
Set rs = oCon.CargaRecordSet(sql)
Set GetMovHabBovAgeCajero = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetMovUserBilletaje(ByVal psOpeCod As String, ByVal psCodUser As String, ByVal pdFecsis As Date, Optional ByVal pbnMovNro As Boolean = False) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function

sql = " SELECT  TOP 1 M.CMOVNRO, M.NMOVNRO, MUE.cEfectivoCod, E.nEfectivoValor, MUE.cUser " _
    & " FROM    MOV M " _
    & "         JOIN MOVUSEREFECTIVO MUE ON MUE.nMOVNRO =M.nMOVNRO " _
    & "         JOIN " & vsBaseComunes & "EFECTIVO E ON  E.cEfectivoCod= MUE.cEfectivoCod " _
    & " WHERE   M.COPECOD='" & psOpeCod & "' AND MUE.cUser ='" & psCodUser & "' and M.nMovFlag in ('" & gMovFlagVigente & "')" _
    & "         And Substring(M.cMovNro,1,8)='" & Format(pdFecsis, "yyyymmdd") & "' "
   
GetMovUserBilletaje = ""
Set rs = oCon.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetMovUserBilletaje = IIf(pbnMovNro, rs!nMovNro, rs!cMovNro)
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion
Set oCon = Nothing

End Function
Public Function GrabaRegistroEfectivo(ByVal psFormatoFecha As String, _
                                    ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                    ByVal rsBillIngreso As ADODB.Recordset, ByVal rsMonIng As ADODB.Recordset, _
                                    ByVal psCodUser As String, ByVal pbNuevoMov As Boolean) As Integer
            
Dim oDMov As DMov
Set oDMov = New DMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
On Error GoTo GrabaRegistroEfectivoErr

GrabaRegistroEfectivo = 1
oDMov.BeginTrans
lbTrans = True
If pbNuevoMov Then
    oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oDMov.GetnMovNro(psMovNro)
Else
    lnMovNro = oDMov.GetnMovNro(psMovNro)
End If
'guardamos la cuenta en el debe
If Not rsBillIngreso Is Nothing Then
    If rsBillIngreso.State = adStateOpen Then
        If Not rsBillIngreso.EOF And Not rsBillIngreso.BOF Then
            Do While Not rsBillIngreso.EOF
                'If rsBillIngreso!Monto > 0 Then
                    If pbNuevoMov Then
                        oDMov.InsertaMovUserEfectivo lnMovNro, psCodUser, rsBillIngreso!cEfectivoCod, rsBillIngreso!Monto
                    Else
                        oDMov.ActualizaMovUserEfectivo lnMovNro, psCodUser, rsBillIngreso!cEfectivoCod, rsBillIngreso!Monto
                    End If
                'End If
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
                'If rsMonIng!Monto > 0 Then
                    If pbNuevoMov Then
                        oDMov.InsertaMovUserEfectivo lnMovNro, psCodUser, rsMonIng!cEfectivoCod, rsMonIng!Monto
                    Else
                        oDMov.ActualizaMovUserEfectivo lnMovNro, psCodUser, rsMonIng!cEfectivoCod, rsMonIng!Monto
                    End If
                'End If
                rsMonIng.MoveNext
            Loop
        End If
        rsMonIng.Close: Set rsMonIng = Nothing
    End If
End If
oDMov.CommitTrans
lbTrans = False

GrabaRegistroEfectivo = 0
Set oDMov = Nothing

Exit Function
GrabaRegistroEfectivoErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaRegistroEfectivo", Err.Description & " Error Registro de Efectivo Cajero"
End Function
Public Function GetBilletajeCajero(ByVal psMovNro As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, Optional ByVal psDenominacion As String = "") As ADODB.Recordset
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

sql = "   SELECT Case " _
    & "         WHEN Substring(E.cEfectivoCod,2,1)='B'  Then 'BILLETAJE' + Convert(char(10),E.nEfectivoValor) " _
    & "         WHEN Substring(E.cEfectivoCod,2,1)='M'  Then 'MONEDA   ' + Convert(char(10),E.nEfectivoValor) " _
    & "         END as Descripcion , " _
    & "         ISNULL(MovEfect.Cantidad,0) as Cantidad , ISNULL(MovEfect.Importe,0) as Monto, " _
    & "         E.cEfectivoCod , E.nEfectivoValor " _
    & " FROM   EFECTIVO E " _
    & "         Left Join " _
    & "                 (SELECT ABS(ISNULL(MOE.NMONTO,0))/E.nEfectivoValor  AS CANTIDAD, " _
    & "                         ABS(ISNULL(MOE.NMONTO,0)) AS IMPORTE, MOE.cEfectivoCod, " _
    & "                         E.nEfectivoValor , m.cOpeCod, MOE.cUser, M.CMOVNRO, M.NMOVNRO, M.nMovFlag " _
    & "                  FROM   MOV M " _
    & "                         JOIN MOVUSEREFECTIVO MOE ON MOE.NMovNro =M.NMovNro " _
    & "                         JOIN EFECTIVO E ON E.cEfectivoCod = MOE.cEfectivoCod " _
    & "                 WHERE   M.CMOVNRO= '" & psMovNro & "' AND MOE.cUser='" & psCodUser & "') AS MovEfect " _
    & "        ON MovEfect.cEfectivoCod = E.cEfectivoCod " _
    & " WHERE E.cEfectivoCod LIKE '" & pnMoneda & "%' and MovEfect.nMovFlag IN ('" & gMovFlagVigente & "') " & lsFiltroDenom & " " _
    & "       AND NOT EXISTS (  SELECT  M1.CMOVNRO " _
    & "                         FROM    MOV M1 JOIN MOVREF MR ON MR.NMOVNRO = M1.NMOVNRO " _
    & "                         WHERE   M1.nMovFlag IN (" & gMovFlagVigente & ") AND MR.NMOVNROREF = MovEfect.NMOVNRO ) " _
    & "       Order BY E.nEfectivoValor Desc"



Set rs = oConect.CargaRecordSet(sql)
Set GetBilletajeCajero = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GrabaDevolucionABoveda(ByVal psFormatoFecha As String, _
                                    ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                    ByVal psCtaDebe As String, ByVal psCtaHaber As String, _
                                    ByVal rsBillIngreso As ADODB.Recordset, ByVal rsMonIng As ADODB.Recordset, _
                                    ByVal psAreaCod As String, ByVal psAgeCod As String, _
                                    ByVal psPersCodUser As String, ByVal pnTotal As Currency, _
                                    ByVal pnMovNroReg As Long) As Integer
            
Dim oDMov As DMov
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New DMov
On Error GoTo GrabaDevolucionABovedaErr

lnMovItem = 0: lnMovOrden = 0

GrabaDevolucionABoveda = 1
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
'guardamos la cuenta en el debe
If Not rsBillIngreso Is Nothing Then
    If rsBillIngreso.State = adStateOpen Then
        If Not rsBillIngreso.EOF And Not rsBillIngreso.BOF Then
            Do While Not rsBillIngreso.EOF
                If rsBillIngreso!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, rsBillIngreso!Monto
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
                    oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, rsMonIng!cEfectivoCod
                End If
                rsMonIng.MoveNext
            Loop
        End If
        rsMonIng.Close: Set rsMonIng = Nothing
    End If
End If
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, psAgeCod, psAreaCod


lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnTotal * -1
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovGasto lnMovNro, psPersCodUser, ""

oDMov.InsertaMovRef lnMovNro, pnMovNroReg

oDMov.CommitTrans
lbTrans = False
GrabaDevolucionABoveda = 0

Set oDMov = Nothing

Exit Function
GrabaDevolucionABovedaErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaDevolucionABoveda", Err.Description & " Error Devolucion de Efectivo de Cajero a Boveda"
End Function
Public Function GetMovUserBilletajeDevuelto(ByVal psOpeCod As String, ByVal psCodUser As String, ByVal pdFecsis As Date) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function

sql = " SELECT  TOP 1 M.CMOVNRO, MUE.cEfectivoCod, E.nEfectivoValor, MUE.cUser " _
    & " FROM    MOV M " _
    & "         JOIN MOVUSEREFECTIVO MUE ON MUE.NMOVNRO =M.NMOVNRO " _
    & "         JOIN " & vsBaseComunes & "EFECTIVO E ON  E.cEfectivoCod= MUE.cEfectivoCod " _
    & "         JOIN (SELECT M1.CMOVNRO , M1.NMOVNRO , MR.NMOVNROREF FROM MOV M1 JOIN MOVREF MR ON MR.NMOVNRO = M1.NMOVNRO " _
    & "               WHERE M1.nMovFlag IN ('" & gMovFlagVigente & "')) AS Dev " _
    & "         ON Dev.NMovnroRef = M.NMovNro " _
    & " WHERE   M.COPECOD='" & psOpeCod & "' AND MUE.cUser ='" & psCodUser & "' " _
    & "         And Substring(M.cMovNro,1,8)='" & Format(pdFecsis, "yyyymmdd") & "'  and M.nMovFlag IN (" & gMovFlagVigente & ")"
   
GetMovUserBilletajeDevuelto = ""
Set rs = oCon.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetMovUserBilletajeDevuelto = rs!cMovNro
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GrabaTransferenciaCajero(ByVal psFormatoFecha As String, _
                                    ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                    ByVal pnImporte As Currency, _
                                    ByVal psCtaContHaber As String, ByVal psPersCodDest As String, _
                                    ByVal psCtaContDebe As String, ByVal psPersCodOrig As String, _
                                    ByVal psUserOrig As String, ByVal psUserDest As String) As Integer
Dim oMov As DMov
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim lbTrans As Boolean
Dim lnMovNro As Long


Set oMov = New DMov
GrabaTransferenciaCajero = 1
On Error GoTo GrabaTransferenciaCajeroErr

oMov.BeginTrans
lbTrans = True
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebe, pnImporte
lnMovOrden = lnMovOrden + 1
oMov.InsertaMovGasto lnMovNro, psPersCodOrig, psPersCodDest

lnMovItem = lnMovItem + 1: lnMovOrden = 0
oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber, pnImporte * -1
lnMovOrden = lnMovOrden + 1

oMov.InsertaMovHabDev lnMovNro, psUserDest, pnImporte
oMov.CommitTrans
lbTrans = False
GrabaTransferenciaCajero = 0
Set oMov = Nothing
Exit Function
GrabaTransferenciaCajeroErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GrabaConfHabilitaAgencia(ByVal psFormatoFecha As String, _
                                        ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                        ByVal psMovNroHab As String, ByVal pnImporte As Currency, _
                                        ByVal psCtaContHaber As String, ByVal psAreaCod As String, ByVal psAgeCod As String, _
                                        ByVal psCtaContDebe As String, ByVal psPersCodUser As String, _
                                        ByVal psUser As String) As Integer
Dim oMov As DMov
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New DMov

oMov.BeginTrans
lbTrans = True
GrabaConfHabilitaAgencia = 1
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovHabDev lnMovNro, psUser, pnImporte
oMov.InsertaMovRef lnMovNro, psMovNroHab

'lnMovItem = lnMovItem + 1: lnMovOrden = 0
'oMov.InsertaMovCta psMovNro, lnMovItem, psCtaContDebe, pnImporte
'lnMovOrden = lnMovOrden + 1
'oMov.InsertaMovObj psMovNro, lnMovItem, lnMovOrden, ObjPersona
'oMov.InsertaMovObjPers psMovNro, lnMovItem, lnMovOrden, psPersCodUser

'lnMovItem = lnMovItem + 1: lnMovOrden = 0
'oMov.InsertaMovCta psMovNro, lnMovItem, psCtaContHaber, pnImporte * -1
'lnMovOrden = lnMovOrden + 1
'oMov.InsertaMovObj psMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
'oMov.InsertaMovObjAgenciaArea psMovNro, lnMovItem, lnMovOrden, psAgeCod, psAreaCod


oMov.CommitTrans
lbTrans = False
GrabaConfHabilitaAgencia = 0
Set oMov = Nothing
Exit Function
GrabaConfHabilitaAgencia:
    If lbTrans Then
    oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetConfHabBovAgencia(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                    ByVal psAreaCod As String, ByVal psAgeCod As String, _
                                    ByVal psCodUser As String) As ADODB.Recordset
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltroDenom As String
Dim lsFiltroUser As String
Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
lsFiltroUser = ""
If psCodUser <> "" Then
    lsFiltroUser = "AND D.CUSER ='" & psCodUser & "' "
End If


sql = "SELECT   CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(MCONF.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2)  as Fecha,          " _
    & "         MCONF.cOpeDesc ,D.CUSER,  D.CPERSNOMBRE, D.IMPORTE, MCONF.cMovDesc , MCONF.NMovNro, D.CPERSCOD " _
    & " FROM    MOV M " _
    & "         JOIN (  SELECT  M.CMOVNRO, M.nMOVNRO, ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)) AS IMPORTE, " _
    & "                         MOP.CPERSCOD , RH.CUSER, P.CPERSNOMBRE, MO.COBJETOCOD " _
    & "                 FROM    MOV M " _
    & "                         JOIN MOVCTA MC  ON MC.nMOVNRO =M.nMOVNRO " _
    & "                         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "                         JOIN MOVGasto MOP ON MOP.nMOVNRO = MO.nMOVNRO " _
    & "                         JOIN RRHH RH ON RH.CPERSCOD = MOP.CPERSCOD " _
    & "                         JOIN " & vsBasePesonas & "PERSONA P ON P.CPERSCOD = MOP.CPERSCOD " _
    & "                  WHERE   M.COPECOD ='" & psOpeCod & "' AND MC.CCTACONTCOD>0 ) AS D " _
    & "         ON D.nMOVNRO = M.nMOVNRO " _
    & "         JOIN (  SELECT  M.CMOVNRO, M.nMOVNRO, MOA.CAREACOD, MOA.CAGECOD, MO.COBJETOCOD " _
    & "                 FROM    MOV M " _
    & "                         JOIN MOVCTA MC  ON MC.nMOVNRO =M.nMOVNRO " _
    & "                         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "                         JOIN MOVOBJ MO ON MO.nMOVNRO = MC.nMOVNRO AND MO.nMOVITEM = MC.nMOVITEM AND MO.COBJETOCOD ='13' " _
    & "                         JOIN MOVOBJAREAAGENCIA MOA ON MOA.nMOVNRO = MO.nMOVNRO AND MOA.nMOVITEM = MO.nMOVITEM AND MOA.nMOVOBJORDEN = MO.nMOVOBJORDEN " _
    & "                 WHERE   M.COPECOD ='" & psOpeCod & "') AS H " _
    & "         ON H.nMOVNRO = M.nMOVNRO "
sql = sql + "   JOIN   (SELECT  M1.CMOVNRO, M1.nMOVNRO,  M1.cMovDesc , MR.nMOVNROREF, O.cOpeDesc " _
    & "                 FROM    MOV M1  " _
    & "                         JOIN MOVREF MR ON MR.nMOVNRO = M1.nMOVNRO " _
    & "                         JOIN " & vsBaseComunes & "OPETPO O ON O.COPECOD = M1.COPECOD " _
    & "           WHERE   M1.nMovFlag IN ('" & gMovFlagVigente & "') )   AS MCONF " _
    & "         ON MCONF.nMOVNROREF = M.nMOVNRO " _
    & " WHERE   M.COPECOD ='" & psOpeCod & "' " _
    & "         AND M.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagModificado & ")" _
    & "         AND H.CAREACOD ='" & psAreaCod & "' AND H.CAGECOD='" & psAgeCod & "' " _
    & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdDesde, "yyyymmdd") & "' " & lsFiltroUser _
    & "         ORDER BY M.CMOVNRO "


Set rs = oConect.CargaRecordSet(sql)
Set GetConfHabBovAgencia = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetDevCajero(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                    ByVal psCodUser As String) As ADODB.Recordset
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltroDenom As String
Dim lsFiltroUser As String

Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
lsFiltroUser = ""
If psCodUser <> "" Then
    lsFiltroUser = " AND MUE.cUser ='" & psCodUser & "' "
End If


sql = "SELECT   CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(DEV.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(DEV.CMOVNRO,9,2)+ ':' + SUBSTRING(DEV.CMOVNRO,11,2) + ':' + SUBSTRING(DEV.CMOVNRO,13,2)  AS FECHA, " _
    & "         Dev.cOpeDesc, MUE.cUser, P.CPERSNOMBRE, SUM(MUE.nMonto) AS MontoDevuelto, " _
    & "         Dev.cMovDesc , DEV.nMovNro " _
    & " FROM    MOV M " _
    & "         JOIN MOVUSEREFECTIVO MUE ON MUE.nMOVNRO =M.nMOVNRO " _
    & "         JOIN RRHH RH ON RH.CUSER=MUE.CUSER " _
    & "         JOIN " & vsBasePesonas & "PERSONA P ON P.CPERSCOD = RH.CPERSCOD " _
    & "         JOIN (  SELECT  M1.CMOVNRO, M1.nMOVNRO  , MR.nMOVNROREF, O.cOpeDesc , M1.cMovDesc " _
    & "                 FROM    MOV M1 JOIN MOVREF MR ON MR.nMOVNRO = M1.nMOVNRO " _
    & "                         JOIN " & vsBaseComunes & "OPETPO O ON O.COPECOD =M1.COPECOD" _
    & "                 WHERE   M1.nMovFlag IN (" & gMovFlagVigente & ")) AS Dev " _
    & "         ON Dev.nMovnroRef = M.nMovNro " _
    & " WHERE   M.COPECOD='" & psOpeCod & "' " & lsFiltroUser _
    & "         And Substring(M.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' and '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "         And M.nMovFlag IN (" & gMovFlagVigente & ") " _
    & " GROUP BY DEV.nMOVNRO, DEV.CMOVNRO, MUE.cUser , P.CPERSNOMBRE  , Dev.cMovDesc ,  Dev.cOpeDesc "

Set rs = oConect.CargaRecordSet(sql)
Set GetDevCajero = rs
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetTransfEntreCajeros(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                    ByVal psCodUser As String) As ADODB.Recordset
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltroDenom As String
Dim lsFiltroUser As String

Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

If psCodUser <> "" Then
    lsFiltroUser = " AND ORIGEN.CUSER='" & psCodUser & "'  "
End If


sql = " SELECT   DEST.FECHA, DEST.COPEDESC, RIGHT(DEST.CMOVNRO,4) , DEST.CPERSNOMBRE , DEST.IMPORTE , " _
    & "          DEST.CMOVDESC , DEST.NMovNro, DEST.CPERSCOD, DEST.CMovNro, DEST.CUSER AS Destino " _
    & " From " _
    & "         (SELECT CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2)  AS FECHA, " _
    & "                 O.COPEDESC , RH.CUSER, P.CPERSNOMBRE ,ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS IMPORTE, " _
    & "                 M.CMOVDESC, M.cMovNro, M.nMovNro, MP.cPersCodDest CPERSCOD " _
    & "          FROM   MOV M " _
    & "                 JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO " _
    & "                 LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND MC.nMOVITEM = ME.nMOVITEM " _
    & "                 JOIN MOVGasto MP ON MP.nMOVNRO = MO.nMOVNRO " _
    & "                 JOIN RRHH   RH ON RH.CPERSCOD = MP.CPERSCODDest " _
    & "                 JOIN " & vsBasePesonas & "PERSONA    P  ON P.CPERSCOD = RH.CPERSCOD " _
    & "                 JOIN " & vsBaseComunes & "OPETPO     O ON O.COPECOD =M.COPECOD " _
    & "         WHERE   M.COPECOD ='" & psOpeCod & "' AND MC.NMOVIMPORTE >0 " _
    & "                 AND nMovFlag IN ('" & gMovFlagVigente & "')  " _
    & "         ) AS DEST " _
    & " Join "
sql = sql + "   (SELECT M.CMOVNRO, M.nMOVNRO, RH.CUSER " _
    & "         FROM    MOV M " _
    & "                 JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO " _
    & "                 JOIN MOVOBJ     MO ON MO.nMOVNRO = MC.nMOVNRO AND MO.nMOVITEM = MC.nMOVITEM " _
    & "                 JOIN MOVGasto MP ON MP.nMOVNRO = MO.nMOVNRO " _
    & "                 JOIN RRHH   RH ON RH.CPERSCOD = MP.CPERSCOD " _
    & "         WHERE   M.COPECOD ='" & psOpeCod & "' AND MC.NMOVIMPORTE <0 " _
    & "         ) AS ORIGEN " _
    & " ON ORIGEN.nMOVNRO = DEST.nMOVNRO " _
    & " WHERE SUBSTRING(DEST.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' " & lsFiltroUser
    
Set rs = oConect.CargaRecordSet(sql)
Set GetTransfEntreCajeros = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GrabaCompraVenta(ByVal psFormatoFecha As String, _
                                 ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                 ByVal pnImporte As Currency, ByVal pnTipoCambio As Currency, _
                                 ByVal psCodPers As String) As Integer
Dim oMov As DMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New DMov

On Error GoTo GrabaCompraVentaErr
oMov.BeginTrans
lbTrans = True
GrabaCompraVenta = 1
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovCompraVenta lnMovNro, psCodPers, pnImporte
oMov.InsertaMovTpoCambio lnMovNro, pnTipoCambio

oMov.CommitTrans
lbTrans = False
GrabaCompraVenta = 0
Set oMov = Nothing
Exit Function
GrabaCompraVentaErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetCompraVenta(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                              ByVal psCodUser As String) As ADODB.Recordset
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String

Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
If psCodUser <> "" Then
    lsFiltroUser = "  AND RIGHT(M.CMOVNRO,4)='" & psCodUser & "' "
End If

sql = " Select  CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2)  AS FECHA ," _
    & "         O.COPEDESC, RIGHT(M.CMOVNRO,4)  AS CUSER , P.CPERSNOMBRE , MC.NMOVIMPORTE, " _
    & "         M.cMovDesc ,  m.NMOVNRO, P.CPERSCOD ,  m.COPECOD, nMovTpoCambio as TipoCambio " _
    & " From    MOV M " _
    & "         JOIN MovCompraVenta MC  ON MC.NMOVNRO =M.NMOVNRO " _
    & "         JOIN MOVTPOCAMBIO MT    ON MT.NMOVNRO = MC.NMOVNRO " _
    & "         JOIN " & vsBasePesonas & "PERSONA P      ON P.CPERSCOD = MC.CPERSCOD " _
    & "         JOIN " & vsBaseComunes & "OPETPO O       ON O.COPECOD = M.COPECOD " _
    & " WHERE   M.NMOVFLAG =" & gMovFlagVigente & " AND M.COPECOD='" & psOpeCod & "' " & lsFiltroUser _
    & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' and  '" & Format(pdHasta, "yyyymmdd") & "' " _
    & " ORDER BY M.NMOVNRO "
 
Set rs = oConect.CargaRecordSet(sql)
Set GetCompraVenta = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function ValidaReciboServ(ByVal psOpeCod As String, ByVal psNroRecibo As String, ByVal psReferencia As String) As Boolean
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String

Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
ValidaReciboServ = False

sql = " SELECT  M.NMOVNRO, MS.cNroDoc " _
    & " FROM    MOV M " _
    & "         JOIN MovOpeVarias MS ON MS.NMOVNRO = M.NMOVNRO " _
    & " WHERE   M.COPECOD ='" & psOpeCod & "' AND MS.cNroDoc ='" & psNroRecibo & "' " _
    & "         AND M.NMOVFLAG =" & gMovFlagVigente & " and MS.cReferencia='" & psReferencia & "' "

Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    ValidaReciboServ = True
End If
rs.Close
Set rs = Nothing
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GrabaCajeroPagoServicios(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                 ByVal pnImporte As Currency, ByVal psNroRecibo As String, ByVal psReferencia As String) As Integer
Dim oMov As DMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New DMov

On Error GoTo GrabaCajeroPagoServiciosErr
oMov.BeginTrans
lbTrans = True
GrabaCajeroPagoServicios = 1
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovOpeVarias lnMovNro, Trim(psNroRecibo), Trim(psReferencia), pnImporte
oMov.CommitTrans
lbTrans = False
GrabaCajeroPagoServicios = 0
Set oMov = Nothing
Exit Function
GrabaCajeroPagoServiciosErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetServicios(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                              ByVal psCodUser As String, ByVal psCodAge As String) As ADODB.Recordset
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String

Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
If psCodUser <> "" Then
    lsFiltroUser = "  AND RIGHT(M.CMOVNRO,4)='" & psCodUser & "' "
End If

sql = " SELECT  CONVERT(CHAR(10),CONVERT(DATETIME, SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) AS  FECHA, " _
    & "         O.COPEDESC, RIGHT(M.CMOVNRO,4) AS CUSER, CONVERT(CHAR(25),MS.cNroDoc) + MS.cReferencia AS cDocumentos, " _
    & "         MS.nMovImporte , M.CMOVDESC,  m.NMOVNRO " _
    & " FROM    MOV M " _
    & "         JOIN MovOpeVarias MS                ON MS.NMOVNRO =M.NMOVNRO " _
    & "         JOIN " & vsBaseComunes & "OPETPO O  ON O.COPECOD =M.COPECOD   " _
    & " WHERE   M.NMOVFLAG =" & gMovFlagVigente & " AND M.COPECOD='" & psOpeCod & "' AND SUBSTRING(M.CMOVNRO,18,2) ='" & psCodAge & "' " & lsFiltroUser _
    & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' and  '" & Format(pdHasta, "yyyymmdd") & "' " _
    & " ORDER BY M.NMOVNRO "
 
Set rs = oConect.CargaRecordSet(sql)
Set GetServicios = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetIngEfectFalt(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                               ByVal psCodUser As String, ByVal psCodAge As String) As ADODB.Recordset
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String

Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
If psCodUser <> "" Then
    lsFiltroUser = "  AND RIGHT(M.CMOVNRO,4)='" & psCodUser & "' "
End If

sql = " SELECT  CONVERT(CHAR(10),CONVERT(DATETIME, SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) AS  FECHA, " _
    & "         O.COPEDESC, RIGHT(M.CMOVNRO,4) AS CUSER, CONVERT(CHAR(25),MS.cNroDoc) + MS.cReferencia AS cDocumentos, " _
    & "         MS.nMovImporte , M.CMOVDESC,  m.NMOVNRO " _
    & " FROM    MOV M " _
    & "         JOIN MovOpeVarias MS                ON MS.NMOVNRO =M.NMOVNRO " _
    & "         JOIN " & vsBaseComunes & "OPETPO O  ON O.COPECOD =M.COPECOD   " _
    & " WHERE   M.NMOVFLAG =" & gMovFlagVigente & " AND M.COPECOD='" & psOpeCod & "' AND SUBSTRING(M.CMOVNRO,18,2) ='" & psCodAge & "' " & lsFiltroUser _
    & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' and  '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "         AND EXISTS (SELECT  M1.NMOVNRO " _
    & "                     FROM    MOV M1 JOIN MOVREF MR ON MR.NMOVNROREF=M1.NMOVNRO " _
    & "                     WHERE   M1.NMOVFLAG =" & gMovFlagVigente & " AND MR.NMOVNRO =M.NMOVNRO ) " _
    & " ORDER BY M.NMOVNRO "
 
Set rs = oConect.CargaRecordSet(sql)
Set GetIngEfectFalt = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetOpeIngEgreCaptaciones(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                                    ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim oConect As DConecta

Set rs = New ADODB.Recordset
If psCodAge <> "" Then
    Select Case pnTipoGrupo
        Case gGruposIngEgreAgeLocal
            lsFiltroAge = " AND  SUBSTRING(MC.cCtaCod,4,2)='" & psCodAge & "'"
        Case gGruposIngEgreOtraAgencia
            lsFiltroAge = " AND  SUBSTRING(MC.cCtaCod,4,2)<>'" & psCodAge & "'"
        Case Else
            lsFiltroAge = ""
    End Select
    
End If
If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.CMOVNRO,4)  = '" & psCodUser & "' "
End If


Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

sql = "SELECT   OPECOL.CCODAGE, cGrupoNombre, COUNT(nTotalMov) AS nTotalMov , " _
    & "         SUM(CASE WHEN G.cIngEgr='I' THEN EFECTIVO ELSE EFECTIVO*-1 END ) AS EFECTIVO, " _
    & "         SUM(CASE WHEN G.cIngEgr='I' THEN CHEQUE ELSE CHEQUE*-1 END ) AS CHEQUE, " _
    & "         SUM(CASE WHEN G.cIngEgr='I' THEN ORDENPAGO ELSE ORDENPAGO*-1 END ) AS ORDENPAGO, G.cGrupoCod " _
    & " FROM    GRUPOOPE G JOIN GRUPOSOPE GO ON GO.cGrupoCod = G.cGrupoCod " _
    & "         JOIN (  SELECT   SUBSTRING(MC.cCtaCod,4,2) AS CCODAGE, M.COPECOD, O.COPEDESC, COUNT(*) AS nTotalMov, " _
    & "                         CASE WHEN MD.NDOCTPO IS NULL                 THEN  SUM(MC.nMonto) ELSE 0 END AS EFECTIVO, " _
    & "                         CASE WHEN MD.NDOCTPO=" & TpoDocCheque & "    THEN  SUM(MC.nMonto) ELSE 0 END AS CHEQUE, " _
    & "                         CASE WHEN MD.NDOCTPO=" & TpoDocOrdenPago & " THEN  SUM(MC.nMonto) ELSE 0 END AS ORDENPAGO " _
    & "                 FROM    MOV M " _
    & "                         JOIN MOVCAP MC      ON  MC.NMOVNRO=M.NMOVNRO " _
    & "                         LEFT JOIN MOVDOC MD ON  MD.NMOVNRO = M.NMOVNRO AND MD.NDOCTPO IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
    & "                         JOIN OPETPO O       ON O.COPECOD = M.COPECOD " _
    & "                 WHERE   SUBSTRING(M.CMOVNRO,1,8) ='" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser & lsFiltroAge _
    & "                         AND SUBSTRING(MC.CCTACOD,9,1)='" & pnMoneda & "'  and M.nMovFlag =" & gMovFlagVigente & " AND SUBSTRING(M.CMOVNRO,18,2) ='" & psCodAge & "' " _
    & "                 GROUP BY  M.COPECOD, O.COPEDESC, MD.NDOCTPO, SUBSTRING(MC.cCtaCod,4,2) )AS OPECOL  " _
    & "         ON OPECOL.COPECOD = GO.cOpeCod " _
    & "         WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%'  " _
    & "         GROUP BY OPECOL.CCODAGE, cGrupoNombre, G.cGrupoCod "
sql = sql + " Union " _
    & " (SELECT DISTINCT CAGECOD,CAGEDESCRIPCION,0,0,0,0,'' " _
    & "  FROM   AGENCIAS AG JOIN " _
    & "         (SELECT SUBSTRING(MC.cCtaCod,4,2) AS CCODAGE " _
    & "          FROM MovCap MC JOIN Mov M ON m.nMovNro = mc.nMovNro " _
    & "          WHERE  SUBSTRING(M.CMOVNRO,1,8) ='" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser & lsFiltroAge _
    & "                 AND SUBSTRING(MC.CCTACOD,9,1)='" & pnMoneda & "'  and M.nMovFlag =" & gMovFlagVigente & " AND  SUBSTRING(M.CMOVNRO,18,2) ='" & psCodAge & "'  " _
    & "          ) Mov ON Mov.cCodAge = Ag.cAgecod) " _
    & " ORDER BY CCODAGE "

Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreCaptaciones = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetOpeIngEgreOtrasOpe(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                                     ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim oConect As DConecta

Set rs = New ADODB.Recordset
If psCodAge <> "" Then
    lsFiltroAge = " AND  SUBSTRING(M.CMOVNRO,18,2) ='" & psCodAge & "'"
End If
If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.CMOVNRO,4)  = '" & psCodUser & "' "
End If

Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

'--Otras Operaciones
sql = "  SELECT   cGrupoNombre, COUNT(nTotalMov) AS nTotalMov, " _
    & "         SUM(CASE WHEN G.cIngEgr='I' THEN EFECTIVO ELSE EFECTIVO*-1 END) AS EFECTIVO, " _
    & "         SUM(CASE WHEN G.cIngEgr='I' THEN CHEQUE ELSE CHEQUE*-1 END) AS CHEQUE, " _
    & "         SUM(CASE WHEN G.cIngEgr='I' THEN ORDENPAGO ELSE ORDENPAGO*-1 END) AS ORDENPAGO, GO.cGrupoCod " _
    & " FROM    GRUPOOPE  AS G JOIN GRUPOSOPE GO ON GO.cGrupoCod=G.cGrupoCod " _
    & "         JOIN (  SELECT  M.COPECOD, O.COPEDESC, COUNT(*) AS nTotalMov, " _
    & "                         CASE WHEN MD.NDOCTPO IS NULL                   THEN  SUM(MV.NMOVIMPORTE) ELSE 0 END AS EFECTIVO, " _
    & "                         CASE WHEN MD.NDOCTPO=" & TpoDocCheque & "      THEN  SUM(MV.NMOVIMPORTE) ELSE 0 END AS CHEQUE, " _
    & "                         CASE WHEN MD.NDOCTPO=" & TpoDocOrdenPago & "  THEN  SUM(MV.NMOVIMPORTE) ELSE 0 END AS ORDENPAGO " _
    & "                 FROM    MOV M " _
    & "                         JOIN MOVOPEVARIAS MV ON MV.NMOVNRO=M.NMOVNRO " _
    & "                         LEFT JOIN MOVDOC MD ON MD.NMOVNRO = M.NMOVNRO AND MD.NDOCTPO IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
    & "                         JOIN OPETPO O ON O.COPECOD = M.COPECOD " _
    & "                 WHERE   SUBSTRING(M.CMOVNRO,1,8) ='" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser _
    & "                         AND M.NMOVFLAG =" & gMovFlagVigente & " AND SUBSTRING(M.cOpeCod,3,1)=" & pnMoneda & " " _
    & "                 GROUP BY  M.COPECOD, O.COPEDESC, MD.NDOCTPO ) AS OPESERV " _
    & "         ON OPESERV.cOpeCod = GO.cOpeCod " _
    & " WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%'  " _
    & " GROUP BY cGrupoNombre, GO.cGrupoCod "

Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreOtrasOpe = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetOpeIngEgreCompraVenta(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                                     ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim oConect As DConecta

Set rs = New ADODB.Recordset
If psCodAge <> "" Then
    lsFiltroAge = " AND  SUBSTRING(M.CMOVNRO,18,2) ='" & psCodAge & "'"
End If
If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.CMOVNRO,4) = '" & psCodUser & "' "
End If

Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

'--COMPRA VENTA ME  '--si el reporte es en dolares se debe mostrar solo el campo EFECDOL
'--por ahora sólo mustro la parte soles puesto que sólo estoy mostrando la consulta en soles
'--PARA PODER REALIZAR EL UNION CON LAS DEMAS CONSULTAS
'-- EL SIGNO NO SE BE DIFERENCIAR POR EL GRUPO OPERACIONES PUESTO QUE YA SE DEDUJO A PARTIR
'-- DEL CODIGO DE LA OPERACION MAS NO DEL CAMPO DEL GRUPOOPERACIONES
'--si el reporte es en soles se debe mostrar solo el campo EFECSOLES"
sql = "  SELECT cGrupoNombre, nTotalMov, " _
    & "         " & IIf(pnMoneda = gMonedaNacional, " EFECSOLES", " EFECDOL ") & " AS EFECTIVO, CHEQUE,ORDENPAGO, GO.cGrupoCod " _
    & "         FROM    GRUPOOPE  AS G JOIN GRUPOSOPE GO ON GO.cGrupoCod=G.cGrupoCod " _
    & "         JOIN (  " _
    & "         SELECT COPECOD, COPEDESC, nTotalMov , EFECSOLES, EFECDOL, 0 AS CHEQUE , 0 AS ORDENPAGO " _
    & "         FROM ( " _
    & "             SELECT  COPECOD, COPEDESC, nTotalMov, " _
    & "                     CASE WHEN COPECOD ='" & gOpeCajeroMECompra & "' THEN EFECTIVOSOLES*-1 ELSE EFECTIVOSOLES END AS EFECSOLES, " _
    & "                     CASE WHEN COPECOD ='" & gOpeCajeroMECompra & "' THEN EFECTIVODOLARES ELSE EFECTIVODOLARES*-1 END AS EFECDOL " _
    & "             From " _
    & "                    (SELECT  M.COPECOD, O.COPEDESC, COUNT(*) AS nTotalMov, " _
    & "                             SUM(MC.NMOVIMPORTE*TC.nMovTpoCambio) AS  EFECTIVOSOLES, " _
    & "                             SUM(MC.NMOVIMPORTE) As EFECTIVODOLARES " _
    & "                     FROM    MOV M " _
    & "                             JOIN MOVCOMPRAVENTA MC ON MC.NMOVNRO=M.NMOVNRO " _
    & "                             JOIN MOVTPOCAMBIO  TC  ON TC.NMOVNRO = M.NMOVNRO " _
    & "                             JOIN OPETPO O ON O.COPECOD = M.COPECOD " _
    & "                     WHERE   SUBSTRING(M.CMOVNRO,1,8) ='" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser _
    & "                             AND M.NMOVFLAG =" & gMovFlagVigente & " " _
    & "     GROUP BY M.COPECOD, O.COPEDESC, TC.nMovTpoCambio) AS MTC ) AS TC) AS MTC " _
    & " ON MTC.cOpeCod = GO.cOpeCod "

Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreCompraVenta = rs
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetOpeIngEgreHabDev(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                                     ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim oConect As DConecta

Set rs = New ADODB.Recordset
If psCodAge <> "" Then
    lsFiltroAge = " AND  SUBSTRING(M.CMOVNRO,18,2) ='" & psCodAge & "'"
End If
If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.CMOVNRO,4) = '" & psCodUser & "' "
End If

Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

sql = "SELECT   cGrupoNombre, GO.cGrupoCod, COUNT(nTotalMov), " _
    & "         SUM(NMONTOHABDEV) AS EFECTIVO, SUM(CHEQUE) AS CHEQUE, SUM(ORDENPAGO) AS ORDENPAGO" _
    & " FROM    GRUPOOPE  AS G JOIN GRUPOSOPE GO ON GO.cGrupoCod=G.cGrupoCod " _
    & "         JOIN ( " _
    & "         SELECT  COPECOD, " _
    & "                 CASE    WHEN cAgeOrig='01301' THEN  COPEDESC + ' ' + CAREADESTINOHAB ELSE COPEDESC + ' ' + CAREAORIGENHAB END AS OPEDESC, " _
    & "                         nTotalMov, " _
    & "                 CASE    WHEN cAgeOrig='01301' THEN  EFECTIVO*-1 ELSE EFECTIVO END AS NMONTOHABDEV, " _
    & "                         0 AS CHEQUE , 0 AS ORDENPAGO " _
    & "                 FROM ( " _
    & "                         SELECT  M.COPECOD,  " _
    & "                                 MH.cAreaDest + ISNULL(MH.cAgeDest,'') AS AREACODDEST, " _
    & "                                 A.CAREADESCRIPCION + ' ' +  ISNULL(AG.CAGEDESCRIPCION,'') AS CAREADESTINOHAB , " _
    & "                                 A1.CAREADESCRIPCION + ' ' +  ISNULL(AG1.CAGEDESCRIPCION,'') AS CAREAORIGENHAB , " _
    & "                                 COUNT(*) AS nTotalMov, SUM(MH.nMovImporte) AS EFECTIVO, " _
    & "                                 O.COPEDESC, ISNULL(MH.cAreaOrig,'') AS cAreaOrig, ISNULL(MH.cAgeOrig,'') AS cAgeOrig " _
    & "                         FROM    MOV M " _
    & "                                 JOIN MovHabilitacion MH ON MH.NMOVNRO=M.NMOVNRO " _
    & "                                 JOIN OPETPO O ON O.COPECOD = M.COPECOD " _
    & "                                 JOIN AREAS A ON A.CAREACOD = MH.cAreaDest " _
    & "                                 LEFT JOIN AGENCIAS AG ON AG.CAGECOD = MH.cAgeDest " _
    & "                                 JOIN AREAS A1 ON A1.CAREACOD = MH.cAreaOrig " _
    & "                                 LEFT JOIN AGENCIAS AG1 ON AG1.CAGECOD = MH.cAgeOrig "
sql = sql + "                   WHERE   SUBSTRING(M.CMOVNRO,1,8) = '20010701' " _
    & "                                 AND RIGHT(M.CMOVNRO,4)  = 'EJRS' AND M.NMOVFLAG =0 " _
    & "                                 AND (cAgeOrig='01' OR MH.cAgeDest='01301')" _
    & "                                 AND EXISTS (SELECT  M.NMOVNRO " _
    & "                                             FROM    MOV M1 JOIN MOVREF MR ON MR.NMOVNRO=M1.NMOVNRO " _
    & "                                             WHERE   MR.NMOVNROREF=M.NMOVNRO AND M1.NMOVFLAG ='0'  ) " _
    & "                 GROUP BY  M.COPECOD, O.COPEDESC, MH.cAreaDest, A.CAREADESCRIPCION,MH.cAgeDest, " _
    & "                 AG.CAGEDESCRIPCION, MH.cAreaOrig, MH.cAgeOrig,A1.CAREADESCRIPCION, AG1.CAGEDESCRIPCION ) AS HABDEV ) AS OPEHD " _
    & "         ON OPEHD.cOpeCod = GO.cOpeCod " _
    & " GROUP BY cGrupoNombre, GO.cGrupoCod  "

Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreHabDev = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetCompraVentaRepo(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                              ByVal psCodUser As String, ByVal psCodAge As String, Optional lbHistorica As Boolean = False) As ADODB.Recordset
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion() = False Then Exit Function
If psCodUser <> "" Then
    lsFiltroUser = "  AND RIGHT(M.CMOVNRO,4)='" & psCodUser & "' "
End If
If psCodAge <> "" Then
    lsFiltroAge = "  AND SUBSTRING(M.CMOVNRO,18,2)='" & Format(psCodAge, "00") & "' "
End If

sql = " SELECT   M.NMOVNRO, P.CPERSNOMBRE, nMovImporte AS Dolares , CONVERT(VARCHAR(12),CONVERT(DATETIME, SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHA ,  " _
    & "         nMovTpoCambio as Cambio, nMovImporte* nMovTpoCambio as Soles, " _
    & "         Right(M.cMovNro,4) as cUser, " _
    & "         (SUBSTRING(M.CMOVNRO,9,2) +':'+ SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) ) AS HORA ,M.COPECOD,  " _
    & "         AG.CAGEDESCRIPCION AS AGENCIA, AG.CAGECOD " _
    & " FROM    MOV M " _
    & "         JOIN MOVCOMPRAVENTA MC ON MC.NMOVNRO=M.NMOVNRO " _
    & "         JOIN PERSONA P ON P.CPERSCOD = MC.CPERSCOD " _
    & "         JOIN MOVTPOCAMBIO MT ON MT.NMOVNRO = M.NMOVNRO " _
    & "         JOIN AGENCIAS AG ON AG.CAGECOD = SUBSTRING(M.CMOVNRO,18,2)   " _
    & " WHERE   M.COPECOD='" & psOpeCod & "'  AND M.NMOVFLAG =" & gMovFlagVigente & "  " _
    & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' and  '" & Format(pdHasta, "yyyymmdd") & "' " & lsFiltroUser & lsFiltroAge _
    & " ORDER BY AG.CAGECOD, M.CMOVNRO  "

Set rs = oConect.CargaRecordSet(sql)
Set GetCompraVentaRepo = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetFaltanteCajero(ByVal psOpeCod As String, ByVal psCodUser As String, ByVal psCodAge As String) As ADODB.Recordset
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset
Dim lsOpeCod As String

Set rs = New ADODB.Recordset
Set oConect = New DConecta

If oConect.AbreConexion = False Then Exit Function
Select Case psOpeCod
    Case gOpeHabCajIngEfectRegulaFaltMN
        lsOpeCod = gOpeHabCajRegSobFaltMN
    Case gOpeHabCajIngEfectRegulaFaltME
        lsOpeCod = gOpeHabCajRegSobFaltME
End Select

sql = "     Select  CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.cMovNro,1,8)),103) AS FECHA," _
        & "         MV.nMovImporte as Total , ISNULL(CUB.nCubierto,0) AS nCubierto , " _
        & "         MV.nMovImporte  - ISNULL(CUB.nCubierto,0) AS Pendiente, M.NMOVNRO " _
        & " From    Mov M " _
        & "         JOIN MovOpeVarias MV ON MV.nMovNro = M.nMovNro " _
        & "         LEFT JOIN(SELECT MR.NMOVNROREF, SUM(MV1.nMovImporte) AS nCubierto " _
        & "                   FROM   MOV M1 " _
        & "                          JOIN MOVOPEVARIAS MV1  ON MV1.NMOVNRO =M1.NMOVNRO " _
        & "                          JOIN MOVREF    MR  ON MR.NMOVNRO  =M1.NMOVNRO " _
        & "                   WHERE  M1.COPECOD IN ('" & psOpeCod & "') " _
        & "                          AND M1.NMOVFLAG=" & gMovFlagVigente & " " _
        & "                   GROUP BY MR.NMOVNROREF )  AS CUB " _
        & "         ON CUB.NMOVNROREF = M.NMOVNRO " _
        & " WHERE   M.COPECOD ='" & lsOpeCod & "'   AND M.NMOVFLAG =" & gMovFlagVigente & " AND MV.NMOVIMPORTE<0 AND " _
        & "         RIGHT(M.CMOVNRO,4) ='" & psCodUser & "' AND SUBSTRING(M.CMOVNRO,18,2) ='" & psCodAge & "'"

Set rs = oConect.CargaRecordSet(sql)
Set GetFaltanteCajero = rs
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GrabaRegularizaFaltante(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                 ByVal pnImporte As Currency, ByVal psReferencia As String, ByVal pnMovNroReg As Long) As Integer
Dim oMov As DMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New DMov

On Error GoTo GrabaRegularizaFaltanteErr
oMov.BeginTrans
lbTrans = True
GrabaRegularizaFaltante = 1
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovOpeVarias lnMovNro, "", Trim(psReferencia), pnImporte
oMov.InsertaMovRef lnMovNro, pnMovNroReg
oMov.CommitTrans
lbTrans = False
GrabaRegularizaFaltante = 0
Set oMov = Nothing
Exit Function
GrabaRegularizaFaltanteErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetLibroCaja(psEmpresa As String, psEmpresaRuc As String, psAgenciaCod As String, psAgencia As String, pdFecha As Date, pnMoneda As Moneda, pgbBitCentral As Boolean) As String
    Dim lsCadena As String
    Dim lnAnchoReporte As Long
    Dim oCaj As DCajero
    Set oCaj = New DCajero
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
        
    Dim lsValor As String * 20
    Dim lsMontoDet As String * 15
    Dim lsMonto As String * 15
    Dim lsMontoAcum As String * 15
    
    Dim lnMargen As Integer
    
    Dim lnMonto As Currency
    Dim lnMontoTotal As Currency
    Dim lnMontoAyer As Currency
    Dim lnMontoEntradas As Currency
    Dim lnMontoSalidas As Currency
    Dim lnMontoEntradasAcum As Currency
    Dim lnMontoSalidasAcum As Currency
    
    Dim lbPrimerSimbolo As Boolean
    
    Dim lnDiasAtras As Integer
    Dim lnDiasAtrasTope As Integer
    Dim lbRSVacio As Boolean
    
    lnMontoEntradasAcum = 0
    lnMontoSalidasAcum = 0
    lsMontoDet = ""
    lnAnchoReporte = 90
    lnMargen = 28
    lnDiasAtras = 0
    lnDiasAtrasTope = 5
    '
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnNegritaON & psEmpresa & Space(lnAnchoReporte - Len(psEmpresa)) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnNegritaON & "RUC : " & psEmpresaRuc & Mid(CentrarCadena(psAgencia, lnAnchoReporte), Len("RUC : " & psEmpresaRuc)) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnNegritaON & CentrarCadena("LIBRO CAJA " & IIf(pnMoneda = gMonedaNacional, "MONEDA NACIONAL", "MONEDA EXTRANJERA"), lnAnchoReporte) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    '''lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnNegritaON & CentrarCadena(IIf(pnMoneda = gMonedaNacional, "(NUEVOS SOLES)", "(DOLARES AMERICANOS)"), lnAnchoReporte) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnNegritaON & CentrarCadena(IIf(pnMoneda = gMonedaNacional, "(" & StrConv(gcPEN_PLURAL, vbUpperCase) & ")", "(DOLARES AMERICANOS)"), lnAnchoReporte) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnNegritaON & Space(lnAnchoReporte - Len(Format(pdFecha, gsFormatoFechaView)) - 8) & "Fecha : " & Format(pdFecha, gsFormatoFechaView) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & String(lnAnchoReporte, "=") & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    
    'If psAgenciaCod = "" Then
    '    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnNegritaON & "CONSOLIDADO" & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    'Else
    '    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnNegritaON & "Ag./Of.Esp.  " & psAgenciaCod & " " & psAgencia & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    'End If
    
    
    'SALDO ANTERIOR
    While lnDiasAtras < lnDiasAtrasTope And lnMontoAyer = 0
        lnDiasAtras = lnDiasAtras + 1
        lnMontoAyer = oCaj.GetValorBilletaje(pdFecha - lnDiasAtras, pnMoneda, psAgenciaCod, pgbBitCentral, True)
    Wend
    
    RSet lsMontoAcum = Format(lnMontoAyer, gsFormatoNumeroView)
    lsMonto = ""
    '''lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnNegritaON & "1.  SALDO DEL DIA ANTERIOR            " & "   " & lsMontoDet & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & lsMonto & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & lsMontoAcum & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnNegritaON & "1.  SALDO DEL DIA ANTERIOR            " & "   " & lsMontoDet & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & lsMonto & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & lsMontoAcum & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    
    'INGRESOS DE CAJA
    lnMontoEntradas = Abs(oCaj.GetHabDev(pdFecha, True, pnMoneda, psAgenciaCod, pgbBitCentral))
    lnMontoEntradasAcum = lnMontoEntradasAcum + lnMontoEntradas
    RSet lsMonto = Format(lnMontoEntradas, gsFormatoNumeroView)
    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnNegritaON & "2.  INGRESOS DE LA FECHA              " & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "    HABILITACIONES RECIBIDAS          " & "   " & lsMontoDet & "   " & lsMonto & oImpresora.gPrnSaltoLinea
    lnMontoEntradas = Round(Abs(oCaj.GetValorMovimiento(pdFecha, True, pnMoneda, psAgenciaCod, pgbBitCentral)), 2) 'ingreso del dia
    lnMontoEntradasAcum = lnMontoEntradasAcum + lnMontoEntradas
    RSet lsMonto = Format(lnMontoEntradas, gsFormatoNumeroView)
    lsCadena = lsCadena & Space(lnMargen) & "    INGRESOS POR OPERACIONES          " & "   " & lsMontoDet & "   " & lsMonto & oImpresora.gPrnSaltoLinea
    lnMontoEntradas = Abs(oCaj.GetSobFal(pdFecha, True, pnMoneda, psAgenciaCod, pgbBitCentral))   'sOBRANTE
    lnMontoEntradasAcum = lnMontoEntradasAcum + lnMontoEntradas
    RSet lsMonto = Format(lnMontoEntradas, gsFormatoNumeroView)
    lsCadena = lsCadena & Space(lnMargen) & "    SOBRANTES                         " & "   " & lsMontoDet & "   " & lsMonto & oImpresora.gPrnSaltoLinea
    lnMontoEntradas = 0 'Otros Ingresos
    lnMontoEntradasAcum = lnMontoEntradasAcum + lnMontoEntradas
    RSet lsMonto = Format(lnMontoEntradas, gsFormatoNumeroView)
    RSet lsMontoAcum = Format(lnMontoEntradasAcum, gsFormatoNumeroView)
    lsCadena = lsCadena & Space(lnMargen) & "    OTROS INGRESOS                    " & "   " & lsMontoDet & "   " & lsMonto & "   " & lsMontoAcum & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    
    'EGRESOS DE CAJA
    lnMontoSalidas = Abs(oCaj.GetHabDev(pdFecha, False, pnMoneda, psAgenciaCod, pgbBitCentral))
    lnMontoSalidasAcum = lnMontoSalidasAcum + lnMontoSalidas
    RSet lsMonto = Format(lnMontoSalidas, gsFormatoNumeroView)
    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnNegritaON & "3.  EGRESOS DE LA FECHA               " & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "    HABILITACIONES ENTREGADAS         " & "   " & lsMontoDet & "   " & lsMonto & oImpresora.gPrnSaltoLinea
    lnMontoSalidas = Round(Abs(oCaj.GetValorMovimiento(pdFecha, False, pnMoneda, psAgenciaCod, pgbBitCentral)), 2) 'Egreso del dia
    lnMontoSalidasAcum = lnMontoSalidasAcum + lnMontoSalidas
    RSet lsMonto = Format(lnMontoSalidas, gsFormatoNumeroView)
    lsCadena = lsCadena & Space(lnMargen) & "    EGRESOS POR OPERACIONES           " & "   " & lsMontoDet & "   " & lsMonto & oImpresora.gPrnSaltoLinea
    lnMontoSalidas = Abs(oCaj.GetSobFal(pdFecha, False, pnMoneda, psAgenciaCod, pgbBitCentral))  'sOBRANTE
    lnMontoSalidasAcum = lnMontoSalidasAcum + lnMontoSalidas
    RSet lsMonto = Format(lnMontoSalidas, gsFormatoNumeroView)
    lsCadena = lsCadena & Space(lnMargen) & "    FALTANTES                         " & "   " & lsMontoDet & "   " & lsMonto & oImpresora.gPrnSaltoLinea
    lnMontoSalidas = 0 'Otros Egresos
    lnMontoSalidasAcum = lnMontoSalidasAcum + lnMontoSalidas
    RSet lsMonto = Format(lnMontoSalidas, gsFormatoNumeroView)
    RSet lsMontoAcum = "(" & Format(lnMontoSalidasAcum, gsFormatoNumeroView)
    lsCadena = lsCadena & Space(lnMargen) & "    OTROS EGRESOS                     " & "   " & lsMontoDet & "   " & lsMonto & "   " & lsMontoAcum & ")" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    
    'SALDO EFECTIVO DEL CAJA
    lsMonto = ""
    lnMontoSalidas = Abs(oCaj.GetHabDev(pdFecha, True, pnMoneda, psAgenciaCod, pgbBitCentral))
    lnMontoSalidasAcum = Abs(lnMontoEntradasAcum) - Abs(lnMontoSalidasAcum) + lnMontoAyer
    RSet lsMontoAcum = Format(lnMontoSalidasAcum, gsFormatoNumeroView)
    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnNegritaON & "4. SALDO DE EFECTIVO DEL DIA           " & "   " & lsMontoDet & "   " & lsMonto & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & lsMontoAcum & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    lsMontoAcum = String(100, "=")
    lsCadena = lsCadena & Space(lnMargen) & "                                       " & "   " & lsMontoDet & "   " & lsMonto & "   " & lsMontoAcum & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    
    lsCadena = lsCadena & Space(lnMargen) & "5. DESCOMPOSICION DEL EFECTIVO        " & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    
    
    lnDiasAtras = -1
    lbRSVacio = True
    While lnDiasAtras < lnDiasAtrasTope And lbRSVacio
        lnDiasAtras = lnDiasAtras + 1
        Set rs = oCaj.GetLibroCajaDatos(pdFecha - lnDiasAtras, pnMoneda, True, psAgenciaCod, pgbBitCentral)
        If Not rs.EOF Then lbRSVacio = False
    Wend
    lbPrimerSimbolo = True
    
    lnMonto = 0
    
    RSet lsValor = "DENOMINACION"
    RSet lsMonto = "IMPORTE"
    lsCadena = lsCadena & Space(lnMargen) & "      " & "     " & lsValor & "          " & lsMonto & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "Billetes" & oImpresora.gPrnSaltoLinea
    While Not rs.EOF
        RSet lsValor = Format(rs!nEfectivoValor, gsFormatoNumeroView)
        RSet lsMonto = Format(rs!Monto, gsFormatoNumeroView)
        
        lnMonto = lnMonto + rs!Monto
        If lbPrimerSimbolo Then
            '''lsCadena = lsCadena & Space(lnMargen) & "De    " & "   " & "  " & lsValor & "       " & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            lsCadena = lsCadena & Space(lnMargen) & "De    " & "   " & "  " & lsValor & "       " & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            lbPrimerSimbolo = False
        Else
            lsCadena = lsCadena & Space(lnMargen) & "De    " & "   " & "  " & lsValor & "        " & "  " & lsMonto & oImpresora.gPrnSaltoLinea
        End If
        rs.MoveNext
    Wend
    lsValor = "SUB TOTAL BILLETES"
    lsMonto = ""
    lnMontoTotal = lnMonto
    RSet lsMontoAcum = Format(lnMonto, gsFormatoNumeroView)
    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnNegritaON & "             " & lsValor & "        " & "   " & lsMonto & lsMontoAcum & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    
    
    lnDiasAtras = -1
    lbRSVacio = True
    While lnDiasAtras < lnDiasAtrasTope And lbRSVacio
        lnDiasAtras = lnDiasAtras + 1
        Set rs = oCaj.GetLibroCajaDatos(pdFecha - lnDiasAtras, pnMoneda, False, psAgenciaCod, pgbBitCentral)
        If Not rs.EOF Then lbRSVacio = False
    Wend
    
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "Monedas" & oImpresora.gPrnSaltoLinea
    lnMonto = 0
    While Not rs.EOF
        RSet lsValor = Format(rs!nEfectivoValor, gsFormatoNumeroView)
        RSet lsMonto = Format(rs!Monto, gsFormatoNumeroView)
        
        lnMonto = lnMonto + rs!Monto
        If lbPrimerSimbolo Then
            '''lsCadena = lsCadena & Space(lnMargen) & "De    " & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & "  " & lsValor & "        " & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            lsCadena = lsCadena & Space(lnMargen) & "De    " & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & "  " & lsValor & "        " & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            lbPrimerSimbolo = False
        Else
            lsCadena = lsCadena & Space(lnMargen) & "De    " & "   " & "  " & lsValor & "        " & "  " & lsMonto & oImpresora.gPrnSaltoLinea
        End If
        rs.MoveNext
    Wend
    lsValor = "SUB TOTAL MONEDAS"
    lsMonto = ""
    
    RSet lsMontoAcum = Format(lnMonto, gsFormatoNumeroView)
    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnNegritaON & "             " & lsValor & "        " & "   " & lsMonto & lsMontoAcum & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lnMontoTotal = lnMontoTotal + lnMonto
    lsValor = "TOTAL EFECTIVO AL CIERRE DEL DIA"
    RSet lsMonto = ""
    RSet lsMontoAcum = Format(lnMontoTotal, gsFormatoNumeroView)
    '''lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnNegritaON & "              TOTAL EFECTIVO AL CIERRE DEL DIA              " & lsMonto & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & lsMontoAcum & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnNegritaON & "              TOTAL EFECTIVO AL CIERRE DEL DIA              " & lsMonto & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & lsMontoAcum & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    lsMontoAcum = String(100, "=")
    lsCadena = lsCadena & Space(lnMargen) & "                                                            " & lsMonto & "   " & lsMontoAcum & oImpresora.gPrnSaltoLinea
    
    
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    
    lsCadena = lsCadena & Space(lnMargen) & "      " & "--------------------------" & "                    " & "--------------------------" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "      " & "      ADMINSITRADOR       " & "                    " & "          CAJERO          " & oImpresora.gPrnSaltoLinea
    
    
    GetLibroCaja = lsCadena
End Function


Public Function GetLibroCaja1(psEmpresa As String, psAgenciaCod As String, psAgencia As String, pdFecha As Date, pnMoneda As Moneda, pgbBitCental As Boolean, Optional pbHistoCtaContb As Boolean = False) As String 'JACA 20111229 se agregó pbHistoCtaContb
    Dim lsCadena As String
    Dim lnAnchoReporte As Long
    Dim oCaj As DCajero
    Set oCaj = New DCajero
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
        
    Dim lsValor As String * 9
    Dim lsMonto As String * 15
    
    Dim lnMonto As Currency
    Dim lnMontoAyer As Currency
    Dim lnMontoEntradas As Currency
    Dim lnMontoSalidas As Currency
    'By Capi 09102008 se corrigio caracteres que no reconoce IBM
    lnAnchoReporte = 80
    lsCadena = ""
    'lsCadena = lsCadena & psEmpresa & Space(lnAnchoReporte - Len(psEmpresa) - 2) & Format(pdFecha, "dd") & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & psEmpresa & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & CentrarCadena("LIBRO CAJA", lnAnchoReporte) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & IIf(pnMoneda = gMonedaNacional, "MONEDA NACIONAL", "MONEDA EXTRAJERA") & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    If psAgenciaCod = "" Then
        lsCadena = lsCadena & "CONSOLIDADO" & oImpresora.gPrnSaltoLinea
    Else
        lsCadena = lsCadena & "Ag./Of.Esp.  " & psAgenciaCod & " " & psAgencia & oImpresora.gPrnSaltoLinea
    End If
    
    lsCadena = lsCadena & "Situación de la Caja del día " & Format(pdFecha, gsFormatoFechaView) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "Billetes" & oImpresora.gPrnSaltoLinea
    
    Set rs = oCaj.GetLibroCajaDatos(pdFecha, pnMoneda, True, psAgenciaCod, True)
    
    lnMonto = 0
    While Not rs.EOF
        RSet lsValor = Format(rs!nEfectivoValor, gsFormatoNumeroView)
        RSet lsMonto = Format(rs!Monto, gsFormatoNumeroView)
        
        lnMonto = lnMonto + rs!Monto
        '''lsCadena = lsCadena & "De    " & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & "  " & lsValor & "        " & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
        lsCadena = lsCadena & "De    " & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & "  " & lsValor & "        " & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
        rs.MoveNext
    Wend
    lsValor = "SUB TOTAL"
    RSet lsMonto = Format(lnMonto, gsFormatoNumeroView)
    '''lsCadena = lsCadena & "           " & lsValor & "        " & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    lsCadena = lsCadena & "           " & lsValor & "        " & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    
    
    Set rs = oCaj.GetLibroCajaDatos(pdFecha, pnMoneda, False, psAgenciaCod, True)
    
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "Monedas" & oImpresora.gPrnSaltoLinea
    
    While Not rs.EOF
        RSet lsValor = Format(rs!nEfectivoValor, gsFormatoNumeroView)
        RSet lsMonto = Format(rs!Monto, gsFormatoNumeroView)
        
        lnMonto = lnMonto + rs!Monto
        '''lsCadena = lsCadena & "De    " & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & "  " & lsValor & "        " & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
        lsCadena = lsCadena & "De    " & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & "  " & lsValor & "        " & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
        rs.MoveNext
    Wend
    lsValor = "TOTAL"
    RSet lsMonto = Format(lnMonto, gsFormatoNumeroView)
    '''lsCadena = lsCadena & "           " & lsValor & "        " & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    lsCadena = lsCadena & "           " & lsValor & "        " & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & CentrarCadena("RESUMEN DEL MOVIMIENTO", lnAnchoReporte) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    
    
    lsValor = "TOTAL"
    RSet lsMonto = Format(lnMonto, gsFormatoNumeroView)
    '''lsCadena = lsCadena & "           " & lsValor & "        " & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    lsCadena = lsCadena & "           " & lsValor & "        " & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    
    
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    'Arreglar cuando se tenga el nuevo plan contable
    'lnMontoAyer = oCaj.GetValorBilletaje(pdFecha - 1, pnMoneda, psAgenciaCod, True, True, pbHistoCtaContb)
    lnMontoAyer = oCaj.GetValorBilletaje(pdFecha - 1, pnMoneda, psAgenciaCod, True, True)
    RSet lsMonto = Format(lnMontoAyer, gsFormatoNumeroView)
    '''lsCadena = lsCadena & "  Saldo día Anterior        " & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    lsCadena = lsCadena & "  Saldo día Anterior        " & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    
    lnMontoEntradas = Abs(oCaj.GetValorMovimiento(pdFecha, True, pnMoneda, psAgenciaCod, True))
    RSet lsMonto = Format(lnMontoEntradas, gsFormatoNumeroView)
    '''lsCadena = lsCadena & "  Total entradas de Caja    " & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    lsCadena = lsCadena & "  Total entradas de Caja    " & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    
    RSet lsMonto = Format(lnMontoEntradas + lnMontoAyer, gsFormatoNumeroView)
    '''lsCadena = lsCadena & "  Sub-Total                 " & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    lsCadena = lsCadena & "  Sub-Total                 " & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    
    lnMontoSalidas = Abs(oCaj.GetValorMovimiento(pdFecha, False, pnMoneda, psAgenciaCod, True))
    RSet lsMonto = Format(lnMontoSalidas, gsFormatoNumeroView)
    '''lsCadena = lsCadena & "  Total salidas de Caja     " & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    lsCadena = lsCadena & "  Total salidas de Caja     " & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    
    RSet lsMonto = Format(lnMontoEntradas + lnMontoAyer - lnMontoSalidas, gsFormatoNumeroView)
    '''lsCadena = lsCadena & "  Saldo del día             " & IIf(pnMoneda = gMonedaNacional, gcMN, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    lsCadena = lsCadena & "  Saldo del día             " & IIf(pnMoneda = gMonedaNacional, gcPEN_SIMBOLO, gcME) & lsMonto & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    GetLibroCaja1 = lsCadena
End Function

