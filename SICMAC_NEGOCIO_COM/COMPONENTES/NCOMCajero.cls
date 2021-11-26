VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCajero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsBaseComunes As String
Dim vsBasePesonas As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Class_Initialize()
Dim oImp As COMDConstSistema.DCOMImpresoras
Set oImp = New COMDConstSistema.DCOMImpresoras
Dim oImpresora As New COMFunciones.FCOMVarImpresion
           
oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
   
Set oImp = Nothing
Dim oIni As COMConecta.DCOMClasIni
Set oIni = New COMConecta.DCOMClasIni
vsBaseComunes = oIni.BaseComunes
vsBasePesonas = oIni.BasePersonas
Dim oCons As COMDConstSistema.NCOMConstSistema
Set oCons = New COMDConstSistema.NCOMConstSistema

'Ya no se debe asignar en esta DLL
'gsCodCMAC = oCons.LeeConstSistema(COMDConstantes.gConstSistCodCMAC)

Set oIni = Nothing
Set oCons = Nothing
End Sub
Public Function GetCierreAgencias(ByVal pdFecha As Date, ByVal sOpeCod As String, ByVal sMovDesc As String, ByVal psCodAge As String) As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta
Dim lsTablaDiaria As String
Dim sql As String
Dim rs As ADODB.Recordset

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

lsTablaDiaria = " MovDiario "

sql = "SELECT   CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(M.cMovNro,1,8)),103) + ' ' + SUBSTRING(M.cMovNro,9,2) + ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) Fecha, " _
    & "         O.cOpeDesc, RIGHT(CMOVNRO,4) AS CUSER, P.cPersNombre, SPACE(3) as cSimbolo, 0.00 as nMovimporte, m.nMovNro, 1 as nMoneda, M.cMovDesc " _
    & " FROM    MOVDIARIO M " _
    & "         JOIN OPETPO O ON O.COPECOD = M.COPECOD " _
    & "         JOIN RRHH RH ON RH.CUSER = RIGHT(CMOVNRO,4) " _
    & "         JOIN PERSONA P ON P.CPERSCOD = RH.CPERSCOD " _
    & " WHERE   M.cOpeCod = '" & sOpeCod & "' AND M.nMovFlag NOT IN (2,3,1,5) " _
    & "         AND SUBSTRING(M.CMOVNRO,18,2)='" & psCodAge & "' and LEFT(M.cMovNro,8) = '" & Format$(pdFecha, "yyyymmdd") & "'"

Set rs = oConect.CargaRecordSet(sql)
Set GetCierreAgencias = rs

oConect.CierraConexion
Set oConect = Nothing

End Function


Public Function GrabaHabilitaAgencia(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
            ByVal pnImporte As Double, ByVal psAreaCod As String, ByVal psAgeCod As String, _
            ByVal nMoneda As Moneda, ByVal sUsuOrig As String, ByVal sUsuDest As String) As Integer

Dim oMOv As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMOv = New COMDMov.DCOMMov

On Error GoTo GrabaHabilitaAgenciaErr
oMOv.BeginTrans
lbTrans = True
GrabaHabilitaAgencia = 1
oMOv.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oMOv.GetnMovNro(psMovNro)
oMOv.InsertaMovHabDev lnMovNro, pnImporte, psAreaCod, psAgeCod, nMoneda, sUsuOrig, sUsuDest
oMOv.CommitTrans
lbTrans = False
GrabaHabilitaAgencia = 0
Set oMOv = Nothing
Exit Function
GrabaHabilitaAgenciaErr:
    If lbTrans Then
        oMOv.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetMovHabBovAgeCajero(ByVal psOpeCod As String, ByVal psAreaCod As String, _
                ByVal psAgeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                Optional ByVal psCodUser As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltro As String, sFiltroArea As String

Set oCon = New COMConecta.DCOMConecta
Set rs = New Recordset
If oCon.AbreConexion = False Then Exit Function
lsFiltro = ""
If psCodUser <> "" Then
    lsFiltro = " AND H.cUsuDest = '" & psCodUser & "' "
End If
sFiltroArea = ""
If psAreaCod <> "" Then
    sFiltroArea = "AND H.cAreaCod = '" & psAreaCod & "' "
End If

sql = "SELECT CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(M.cMovNro,1,8)),103) + ' ' " _
    & "+ SUBSTRING(M.cMovNro,9,2) + ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) Fecha, " _
    & "O.cOpeDesc, H.cUsuOrig, ISNULL(RH.cPersNombre,'BOVEDA') cPersNombre, cSimbolo = CASE nMoneda WHEN 1 THEN 'S/.' ELSE 'US$' END, " _
    & "H.nMovImporte, M.nMovNro, H.nMoneda, M.cMovDesc, H.cUsuDest FROM OpeTpo O INNER JOIN Mov M INNER JOIN MovHabDevCajero H " _
    & "LEFT JOIN (Select RH.cUser, P.cPersCod, P.cPersNombre From RRHH RH INNER JOIN Persona P ON RH.cPersCod = P.cPersCod) RH ON H.cUsuOrig = RH.cUser ON M.nMovNro = H.nMovNro " _
    & "ON O.cOpeCod = M.cOpeCod LEFT JOIN (Select MR.nMovNroRef From MovRef MR INNER JOIN Mov M1 ON MR.nMovNro = M1.nMovNro " _
    & "Where M1.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagModificado & ")) MR " _
    & "ON M.nMovNro = MR.nMovNroRef WHERE M.cOpeCod IN ('" & psOpeCod & "') AND " _
    & "M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagModificado & ") " _
    & sFiltroArea & " AND H.cAgeCod = '" & psAgeCod & "' " _
    & "AND SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format$(pdDesde, "yyyymmdd") & "' AND '" & Format$(pdHasta, "yyyymmdd") & "' " & lsFiltro & " " _
    & "And MR.nMovNroRef IS NULL ORDER BY M.cMovNro"
 
Set rs = oCon.CargaRecordSet(sql)
Set GetMovHabBovAgeCajero = rs
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function GetMovRegSobFalt(ByVal psOpeCod As String, ByVal psAgeCod As String, _
                ByVal pdDesde As Date, ByVal pdHasta As Date, _
                Optional ByVal psCodUser As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltro As String

Set oCon = New COMConecta.DCOMConecta
Set rs = New Recordset
If oCon.AbreConexion = False Then Exit Function
lsFiltro = ""
If psCodUser <> "" Then
    lsFiltro = " AND RIGHT(M.cMovNro,4) = '" & psCodUser & "' "
End If

sql = "SELECT CONVERT(Varchar(12),CONVERT(DATETIME,SUBSTRING(M.cMovNro,1,8)),103) + ' ' " _
    & "+ SUBSTRING(M.cMovNro,9,2) + ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) Fecha, " _
    & "O.cOpeDesc, RH.cUser, ISNULL(RH.cPersNombre,'BOVEDA') cPersNombre, cSimbolo = CASE H.nMoneda WHEN 1 THEN 'S/.' ELSE 'US$' END, " _
    & "H.nMovImporte, M.nMovNro, H.nMoneda, M.cMovDesc FROM OpeTpo O JOIN Mov M INNER JOIN MovOpeVarias H " _
    & "ON M.nMovNro = H.nMovNro ON O.cOpeCod = M.cOpeCod " _
    & "LEFT JOIN (Select RH.cUser, P.cPersCod, P.cPersNombre From RRHH RH JOIN Persona P ON RH.cPersCod = P.cPersCod) RH ON RIGHT(M.cMovNro,4) = RH.cUser " _
    & "WHERE M.cOpeCod IN ('" & psOpeCod & "') AND " _
    & "M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagModificado & ") " _
    & "AND SUBSTRING(M.cMovNro,18,2) = '" & psAgeCod & "' " _
    & "AND SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format$(pdDesde, "yyyymmdd") & "' AND '" & Format$(pdHasta, "yyyymmdd") & "' " & lsFiltro & " " _
    & "ORDER BY M.cMovNro"
 
Set rs = oCon.CargaRecordSet(sql)
Set GetMovRegSobFalt = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetBuscarMov(ByVal psOpeCod As String, ByVal psCodUser As String, ByVal psFecha As String, ByVal pnMoneda As Integer, ByVal psCodAge As String) As Boolean
Dim sSQL As String, rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New Recordset
If oCon.AbreConexion = False Then Exit Function

GetBuscarMov = False

    sSQL = "SELECT *    FROM MOVHABDEVCAJERO MH "
    sSQL = sSQL & "     JOIN MOV M ON M.NMOVNRO=MH.NMOVNRO "
    sSQL = sSQL & "     WHERE CusuORIG='" & psCodUser & "' and MH.cAgeCod='" & psCodAge & "' "
    sSQL = sSQL & "  AND M.CMOVNRO LIKE '" & psFecha & "%" & psCodUser & "' and M.copecod='" & psOpeCod & "' and m.nmovflag=0 and mh.nmoneda=" & pnMoneda
    Set rs = oCon.CargaRecordSet(sSQL)
    If Not rs.EOF Then
        GetBuscarMov = True
    End If
    
    Set rs = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
    

End Function


Public Function GetMovHabCajDevABove(ByVal psOpeCod As String, ByVal psAreaCod As String, _
                ByVal psAgeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltro As String

Set oCon = New COMConecta.DCOMConecta
Set rs = New Recordset
If oCon.AbreConexion = False Then Exit Function
lsFiltro = ""

'Filtrar solo de Operaciones

    
'If psAreaCod = "067" Then
'   psAreaCod = "067"
'Else
    psAreaCod = "026"
'End If

sql = "SELECT CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(M.cMovNro,1,8)),103) + ' ' " _
    & "+ SUBSTRING(M.cMovNro,9,2) + ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) Fecha, " _
    & "O.cOpeDesc, RH.cUser, P.cPersNombre, cSimbolo = CASE nMoneda WHEN 1 THEN 'S/.' ELSE 'US$' END, " _
    & "H.nMovImporte, M.nMovNro, H.nMoneda, M.cMovDesc FROM OpeTpo O INNER JOIN Mov M INNER JOIN MovHabDevCajero H " _
    & "ON M.nMovNro = H.nMovNro ON O.cOpeCod = M.cOpeCod INNER JOIN RRHH RH INNER JOIN Persona P ON " _
    & "RH.cPersCod = P.cPersCod ON RIGHT(M.cMovNro,4) = RH.cUser " _
    & "LEFT JOIN (Select MR.nMovNroRef From MovRef MR INNER JOIN Mov M1 ON MR.nMovNro = M1.nMovNro " _
    & "Where M1.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagModificado & ")) MR " _
    & "ON M.nMovNro = MR.nMovNroRef WHERE M.cOpeCod IN ('" & psOpeCod & "') AND " _
    & "M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagModificado & ") " _
    & "AND case when H.cAreaCod<>'026' then '026' else H.cAreaCod end  =  '" & psAreaCod & "' AND H.cAgeCod = '" & psAgeCod & "' " _
    & "AND SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format$(pdDesde, "yyyymmdd") & "' AND '" & Format$(pdHasta, "yyyymmdd") & "' " & lsFiltro & " " _
    & "And MR.nMovNroRef IS NULL ORDER BY M.cMovNro"
 
Set rs = oCon.CargaRecordSet(sql)
Set GetMovHabCajDevABove = rs
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function GetMovUserBilletaje(ByVal psOpeCod As String, ByVal psCodUser As String, _
        ByVal pdFecSis As Date, Optional ByVal pbnMovNro As Boolean = False, _
        Optional ByVal psCodAge As String = "") As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function

sql = "SELECT TOP 1 M.CMOVNRO, M.NMOVNRO, MUE.cEfectivoCod, E.nEfectivoValor, MUE.cUser " _
    & "FROM MOV M JOIN MOVUSEREFECTIVO MUE ON MUE.nMOVNRO = M.nMOVNRO " _
    & "JOIN EFECTIVO E ON  E.cEfectivoCod= MUE.cEfectivoCod " _
    & "WHERE M.COPECOD = '" & psOpeCod & "' AND MUE.cUser ='" & psCodUser & "' and M.nMovFlag in ('" & gMovFlagVigente & "')" _
    & "And Substring(M.cMovNro,1,8)='" & Format(pdFecSis, "yyyymmdd") & "' AND " _
    & "Substring(M.cMovNro,18,2) = '" & psCodAge & "'"
   
GetMovUserBilletaje = ""
Set rs = oCon.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetMovUserBilletaje = IIf(pbnMovNro, rs!nMovNro, rs!cMovnro)
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion
Set oCon = Nothing

End Function

Public Function GrabaRegistroEfectivo(ByVal psFormatoFecha As String, _
            ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
            ByVal rsBillIngMN As ADODB.Recordset, ByVal rsBillIngME As ADODB.Recordset, _
            ByVal rsMonIngMN As ADODB.Recordset, ByVal rsMonIngME As ADODB.Recordset, _
            ByVal psCodUser As String, ByVal pbNuevoMov As Boolean) As Integer
            

Dim oMOv As COMDMov.DCOMMov
Set oMOv = New COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
On Error GoTo GrabaRegistroEfectivoErr

GrabaRegistroEfectivo = 1
oMOv.BeginTrans
lbTrans = True
If pbNuevoMov Then
    oMOv.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oMOv.GetnMovNro(psMovNro)
Else
    lnMovNro = oMOv.GetnMovNro(psMovNro)
End If

'Moneda Nacional
If Not rsBillIngMN Is Nothing Then 'Billetes
    If rsBillIngMN.State = adStateOpen Then
        If Not rsBillIngMN.EOF And Not rsBillIngMN.BOF Then
            Do While Not rsBillIngMN.EOF
                If pbNuevoMov Then
                    oMOv.InsertaMovUserEfectivo lnMovNro, psCodUser, rsBillIngMN("cEfectivoCod"), rsBillIngMN("Monto")
                Else
                    oMOv.ActualizaMovUserEfectivo lnMovNro, psCodUser, rsBillIngMN("cEfectivoCod"), rsBillIngMN("Monto")
                End If
                rsBillIngMN.MoveNext
            Loop
        End If
        rsBillIngMN.Close: Set rsBillIngMN = Nothing
    End If
End If

If Not rsMonIngMN Is Nothing Then 'Monedas
    If rsMonIngMN.State = adStateOpen Then
        If Not rsMonIngMN.EOF And Not rsMonIngMN.BOF Then
            Do While Not rsMonIngMN.EOF
                If pbNuevoMov Then
                    oMOv.InsertaMovUserEfectivo lnMovNro, psCodUser, rsMonIngMN("cEfectivoCod"), rsMonIngMN("Monto")
                Else
                    oMOv.ActualizaMovUserEfectivo lnMovNro, psCodUser, rsMonIngMN("cEfectivoCod"), rsMonIngMN("Monto")
                End If
                rsMonIngMN.MoveNext
            Loop
        End If
        rsMonIngMN.Close: Set rsMonIngMN = Nothing
    End If
End If

'Moneda Extranjera
If Not rsBillIngME Is Nothing Then 'Billetes
    If rsBillIngME.State = adStateOpen Then
        If Not rsBillIngME.EOF And Not rsBillIngME.BOF Then
            Do While Not rsBillIngME.EOF
                If pbNuevoMov Then
                    oMOv.InsertaMovUserEfectivo lnMovNro, psCodUser, rsBillIngME("cEfectivoCod"), rsBillIngME("Monto")
                Else
                    oMOv.ActualizaMovUserEfectivo lnMovNro, psCodUser, rsBillIngME("cEfectivoCod"), rsBillIngME("Monto")
                End If
                rsBillIngME.MoveNext
            Loop
        End If
        rsBillIngME.Close: Set rsBillIngME = Nothing
    End If
    
End If
If Not rsMonIngME Is Nothing Then 'Monedas
    If rsMonIngME.State = adStateOpen Then
        If Not rsMonIngME.EOF And Not rsMonIngME.BOF Then
            Do While Not rsMonIngME.EOF
                If pbNuevoMov Then
                    oMOv.InsertaMovUserEfectivo lnMovNro, psCodUser, rsMonIngME("cEfectivoCod"), rsMonIngME("Monto")
                Else
                    oMOv.ActualizaMovUserEfectivo lnMovNro, psCodUser, rsMonIngME("cEfectivoCod"), rsMonIngME("Monto")
                End If
                rsMonIngME.MoveNext
            Loop
        End If
        rsMonIngME.Close: Set rsMonIngME = Nothing
    End If
End If
oMOv.CommitTrans
lbTrans = False

GrabaRegistroEfectivo = 0
Set oMOv = Nothing

Exit Function
GrabaRegistroEfectivoErr:
    If lbTrans Then
        oMOv.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaRegistroEfectivo", Err.Description & " Error Registro de Efectivo Cajero"
End Function

Public Function GetBilletajeCajero(ByVal psMovNro As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, Optional ByVal psDenominacion As String = "") As ADODB.Recordset
Dim sql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim lsFiltroDenom As String

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
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
            
Dim oMOv As COMDMov.DCOMMov
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMOv = New COMDMov.DCOMMov
On Error GoTo GrabaDevolucionABovedaErr

lnMovItem = 0: lnMovOrden = 0

GrabaDevolucionABoveda = 1
oMOv.BeginTrans
lbTrans = True
oMOv.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oMOv.GetnMovNro(psMovNro)
'guardamos la cuenta en el debe
If Not rsBillIngreso Is Nothing Then
    If rsBillIngreso.State = adStateOpen Then
        If Not rsBillIngreso.EOF And Not rsBillIngreso.BOF Then
            Do While Not rsBillIngreso.EOF
                If rsBillIngreso!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    oMOv.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, rsBillIngreso!Monto
                    lnMovOrden = lnMovOrden + 1
                    oMOv.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oMOv.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, rsBillIngreso!cEfectivoCod
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
                    oMOv.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, rsMonIng!Monto
                    lnMovOrden = lnMovOrden + 1
                    oMOv.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oMOv.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, rsMonIng!cEfectivoCod
                End If
                rsMonIng.MoveNext
            Loop
        End If
        rsMonIng.Close: Set rsMonIng = Nothing
    End If
End If
lnMovOrden = lnMovOrden + 1
oMOv.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
oMOv.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, psAgeCod, psAreaCod


lnMovItem = lnMovItem + 1: lnMovOrden = 0
oMOv.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnTotal * -1
lnMovOrden = lnMovOrden + 1
oMOv.InsertaMovGasto lnMovNro, psPersCodUser, ""

oMOv.InsertaMovRef lnMovNro, pnMovNroReg

oMOv.CommitTrans
lbTrans = False
GrabaDevolucionABoveda = 0

Set oMOv = Nothing

Exit Function
GrabaDevolucionABovedaErr:
    If lbTrans Then
        oMOv.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaDevolucionABoveda", Err.Description & " Error Devolucion de Efectivo de Cajero a Boveda"
End Function
Public Function GetMovUserBilletajeDevuelto(ByVal psOpeCod As String, ByVal psCodUser As String, ByVal pdFecSis As Date) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
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
    & "         And Substring(M.cMovNro,1,8)='" & Format(pdFecSis, "yyyymmdd") & "'  and M.nMovFlag IN (" & gMovFlagVigente & ")"
   
GetMovUserBilletajeDevuelto = ""
Set rs = oCon.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetMovUserBilletajeDevuelto = rs!cMovnro
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

Dim oMOv As COMDMov.DCOMMov
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim lbTrans As Boolean
Dim lnMovNro As Long


Set oMOv = New COMDMov.DCOMMov
GrabaTransferenciaCajero = 1
On Error GoTo GrabaTransferenciaCajeroErr

oMOv.BeginTrans
lbTrans = True
oMOv.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oMOv.GetnMovNro(psMovNro)
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oMOv.InsertaMovCta lnMovNro, lnMovItem, psCtaContDebe, pnImporte
lnMovOrden = lnMovOrden + 1
oMOv.InsertaMovGasto lnMovNro, psPersCodOrig, psPersCodDest

lnMovItem = lnMovItem + 1: lnMovOrden = 0
oMOv.InsertaMovCta lnMovNro, lnMovItem, psCtaContHaber, pnImporte * -1
lnMovOrden = lnMovOrden + 1
 
'oMov.InsertaMovHabDev lnMovNro, psUserDest, pnImporte

oMOv.CommitTrans
lbTrans = False
GrabaTransferenciaCajero = 0
Set oMOv = Nothing
Exit Function
GrabaTransferenciaCajeroErr:
    If lbTrans Then
        oMOv.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GrabaConfHabilitaAgencia(ByVal sMovNro As String, ByVal sOpeCod As String, ByVal sMovDesc As String, _
            ByVal nMovNroHab As Long) As Integer

Dim oMOv As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim nMovNro As Long

Set oMOv = New COMDMov.DCOMMov
oMOv.BeginTrans
lbTrans = True
GrabaConfHabilitaAgencia = 1
oMOv.InsertaMov sMovNro, sOpeCod, sMovDesc, gMovEstContabNoContable, gMovFlagVigente
nMovNro = oMOv.GetnMovNro(sMovNro)
oMOv.InsertaMovRef nMovNro, nMovNroHab
oMOv.CommitTrans
lbTrans = False
GrabaConfHabilitaAgencia = 0
Set oMOv = Nothing
Exit Function
GrabaConfHabilitaAgencia:
    If lbTrans Then
    oMOv.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetConfHabBovAgencia(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                    ByVal psAreaCod As String, ByVal psAgeCod As String, _
                                    ByVal psCodUser As String) As ADODB.Recordset
Dim sql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim lsFiltroDenom As String
Dim lsFiltroUser As String
Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
lsFiltroUser = ""
If psCodUser <> "" Then
    lsFiltroUser = "AND RH.CUSER ='" & psCodUser & "' "
End If

'Filtrar solo de Operaciones
'psAreaCod = "013"
psAreaCod = "026"

sql = "SELECT CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(M1.cMovNro,1,8)),103) + ' ' " _
    & "+ SUBSTRING(M1.cMovNro,9,2) + ':' + SUBSTRING(M1.cMovNro,11,2) + ':' + SUBSTRING(M1.cMovNro,13,2) Fecha, " _
    & "O.cOpeDesc, RH.cUser, P.cPersNombre, cSimbolo = CASE nMoneda WHEN 1 THEN 'S/.' ELSE 'US$' END, " _
    & "H.nMovImporte, M1.nMovNro, H.nMoneda, M.cMovDesc FROM OpeTpo O INNER JOIN Mov M INNER JOIN MovHabDevCajero H " _
    & "INNER JOIN RRHH RH INNER JOIN Persona P ON RH.cPersCod = P.cPersCod ON H.cUsuDEst = RH.cUser ON M.nMovNro = H.nMovNro " _
    & "ON O.cOpeCod = M.cOpeCod INNER JOIN MovRef MR INNER JOIN Mov M1 ON MR.nMovNro = M1.nMovNro ON M.nMovNro = MR.nMovNroRef " _
    & "WHERE M1.cOpeCod IN ('" & psOpeCod & "') AND " _
    & "M1.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagModificado & ") " _
    & "AND H.cAreaCod = '" & psAreaCod & "' AND H.cAgeCod = '" & psAgeCod & "' " _
    & "AND SUBSTRING(M1.cMovNro,1,8) BETWEEN '" & Format$(pdDesde, "yyyymmdd") & "' AND '" & Format$(pdHasta, "yyyymmdd") & "' " & lsFiltroUser & " " _
    & "ORDER BY M.cMovNro"

Set rs = oConect.CargaRecordSet(sql)
Set GetConfHabBovAgencia = rs
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetDevCajero(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                    ByVal psCodUser As String) As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim lsFiltroDenom As String
Dim lsFiltroUser As String
Dim sSQL As String
Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
lsFiltroUser = ""
If psCodUser <> "" Then
    lsFiltroUser = " AND MH.cUsuOrig = '" & psCodUser & "' "
End If

sSQL = "Select CONVERT(Varchar(12),CONVERT(DATETIME,SUBSTRING(M.cMovNro,1,8)),103) + ' ' + " _
    & "SUBSTRING(M.cMovNro,9,2)+ ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) AS Fecha, " _
    & "O.cOpeDesc, MH.cUsuOrig, ISNULL(P.cPersNombre,'BOVEDA') cPersNombre, " _
    & "Moneda = CASE WHEN MH.nMoneda = 1 THEN 'S/.' ELSE 'US$' END, MH.nMovImporte, M.nMovNro, MH.nMoneda, M.cMovDesc " _
    & "FROM OpeTpo O JOIN Mov M JOIN MovHabDevCajero MH LEFT JOIN RRHH RH JOIN Persona P ON RH.cPersCod = P.cPersCod ON " _
    & "MH.cUsuOrig = RH.cUser ON M.nMovNro = MH.nMovNro ON O.cOpeCod = M.cOpeCod " _
    & "WHERE M.cOpeCod = '" & psOpeCod & "' " & lsFiltroUser & " And Substring(M.cMovNro,1,8) " _
    & "BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' And '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "And M.nMovFlag IN (" & gMovFlagVigente & ") " _
    & "Order by M.nMovNro"

Set rs = oConect.CargaRecordSet(sSQL)
Set GetDevCajero = rs
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetTransfEntreCajeros(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                    ByVal psCodUser As String) As ADODB.Recordset
Dim sql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim lsFiltroDenom As String
Dim lsFiltroUser As String

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
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
            ByVal pnImporte As Double, ByVal pnTipoCambio As Double, _
            ByVal psCodPers As String, Optional bLavDinero As Boolean = False, Optional SPersBen As String = "") As Integer
            
Dim oMOv As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMOv = New COMDMov.DCOMMov

On Error GoTo GrabaCompraVentaErr
oMOv.BeginTrans
lbTrans = True
GrabaCompraVenta = 1
oMOv.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oMOv.GetnMovNro(psMovNro)
oMOv.InsertaMovCompraVenta lnMovNro, psCodPers, pnImporte
oMOv.InsertaMovTpoCambio lnMovNro, pnTipoCambio

If SPersBen = "" Then SPersBen = psCodPers

If bLavDinero Then
    oMOv.InsertaMovLavDinero lnMovNro, psCodPers, SPersBen
End If
oMOv.CommitTrans
lbTrans = False
GrabaCompraVenta = 0
Set oMOv = Nothing
Exit Function
GrabaCompraVentaErr:
    If lbTrans Then
        oMOv.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetCompraVenta(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                              ByVal psCodUser As String) As ADODB.Recordset
Dim sql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
If psCodUser <> "" Then
    lsFiltroUser = "  AND RIGHT(M.CMOVNRO,4)='" & psCodUser & "' "
End If

sql = " Select CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) AS FECHA, " _
    & "         O.COPEDESC, RIGHT(M.CMOVNRO,4) AS CUSER, P.CPERSNOMBRE, cSimbolo = 'US$', MC.NMOVIMPORTE, " _
    & "         M.nMovNro, nMoneda = 2,  M.cMovDesc " _
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
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
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
Dim oMOv As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMOv = New COMDMov.DCOMMov

On Error GoTo GrabaCajeroPagoServiciosErr
oMOv.BeginTrans
lbTrans = True
GrabaCajeroPagoServicios = 1
oMOv.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oMOv.GetnMovNro(psMovNro)
oMOv.InsertaMovOpeVarias lnMovNro, Trim(psNroRecibo), Trim(psReferencia), pnImporte
oMOv.CommitTrans
lbTrans = False
GrabaCajeroPagoServicios = 0
Set oMOv = Nothing
Exit Function
GrabaCajeroPagoServiciosErr:
    If lbTrans Then
        oMOv.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetServicios(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                              ByVal psCodUser As String, ByVal psCodAge As String) As ADODB.Recordset
Dim sSQL As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
If psCodUser <> "" Then
    lsFiltroUser = "  AND RIGHT(M.cMovNro,4)='" & psCodUser & "' "
End If

sSQL = "SELECT CONVERT(Varchar(10),CONVERT(DATETIME, SUBSTRING(M.cMovNro,1,8)),103) + ' ' + " _
    & "SUBSTRING(M.cMovNro,9,2) + ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) Fecha, " _
    & "O.cOpeDesc, RIGHT(M.cMovNro,4) Usu, P.cPersNombre, cSimbolo = 'S/.', MS.nMonto, M.nMovNro, " _
    & "nMoneda = 1, M.cMovDesc FROM OpeTpo O JOIN Mov M JOIN MovServicios MS ON MS.nMovNro = M.nMovNro " _
    & "ON O.cOpeCod = M.cOpeCod JOIN RRHH RH JOIN Persona P ON RH.cPersCod = P.cPersCod ON " _
    & "RIGHT(M.cMovNro,4) = RH.cUser WHERE M.nMovFlag = " & gMovFlagVigente & " AND " _
    & "M.cOpeCod = '" & psOpeCod & "' And SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " & lsFiltroUser _
    & "AND SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' And '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "ORDER BY M.nMovNro "
 
Set rs = oConect.CargaRecordSet(sSQL)
Set GetServicios = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetMovFideicomiso(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                              ByVal psCodUser As String, ByVal psCodAge As String) As ADODB.Recordset
Dim sSQL As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
If psCodUser <> "" Then
    lsFiltroUser = "  AND RIGHT(M.cMovNro,4)='" & psCodUser & "' "
End If

sSQL = "SELECT CONVERT(Varchar(10),CONVERT(DATETIME, SUBSTRING(M.cMovNro,1,8)),103) + ' ' + " _
    & "SUBSTRING(M.cMovNro,9,2) + ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) Fecha, " _
    & "O.cOpeDesc, RIGHT(M.cMovNro,4) Usu, F.cCtaCod, cSimbolo = 'S/.', MS.nMonto, M.nMovNro, " _
    & "nMoneda = 1, M.cMovDesc FROM OpeTpo O JOIN Mov M JOIN MovServicios MS ON MS.nMovNro = M.nMovNro " _
    & "ON O.cOpeCod = M.cOpeCod INNER JOIN FIDEICOMISOMOV FM ON FM.nNroMov = MS.nMovNro " _
    & " INNER JOIN FIDEICOMISO F ON FM.cCtaCod = F.cCtaCod WHERE M.nMovFlag = " & gMovFlagVigente & " AND " _
    & "M.cOpeCod = '" & psOpeCod & "' And SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " & lsFiltroUser _
    & "AND SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' And '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "ORDER BY M.nMovNro "
 
Set rs = oConect.CargaRecordSet(sSQL)
Set GetMovFideicomiso = rs
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetIngEfectFalt(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
            ByVal psCodUser As String, ByVal psCodAge As String) As ADODB.Recordset
Dim sSQL As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.cMovNro,4) = '" & psCodUser & "' "
End If

sSQL = "SELECT CONVERT(CHAR(10),CONVERT(DATETIME, SUBSTRING(M.cMovNro,1,8)),103) + ' ' + SUBSTRING(M.cMovNro,9,2)+ ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) AS  FECHA, " _
    & "O.cOpeDesc, RIGHT(M.cMovNro,4) AS cUser, CONVERT(CHAR(25),MS.cNroDoc) + MS.cReferencia AS cDocumentos, " _
    & "cSimbolo = CASE WHEN MS.nMoneda = " & gMonedaNacional & " THEN 'S/.' ELSE 'US$' END, MS.nMovImporte, M.nMovNro, MS.nMoneda, M.cMovDesc FROM MOV M JOIN MovOpeVarias MS ON MS.nMovNro = M.nMovNro " _
    & "JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod WHERE M.nMovFlag = " & gMovFlagVigente & " AND " _
    & "M.cOpeCod = '" & psOpeCod & "' AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " & lsFiltroUser _
    & "AND SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format$(pdDesde, "yyyymmdd") & "' And '" & Format$(pdHasta, "yyyymmdd") & "' " _
    & "AND EXISTS (SELECT  M1.nMovNro FROM MOV M1 JOIN MOVREF MR ON MR.nMovNroREF = M1.nMovNro " _
    & "WHERE M1.nMovFlag = " & gMovFlagVigente & " AND MR.nMovNro = M.nMovNro) " _
    & "ORDER BY M.nMovNro "
 
Set rs = oConect.CargaRecordSet(sSQL)
Set GetIngEfectFalt = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
'
'Public Function GetOpeIngEgreCaptaciones(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
'                                    ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, _
'                                    Optional ByVal psCodCmac As String) As ADODB.Recordset
'Dim ssql As String
'Dim Rs As ADODB.Recordset
'Dim lsFiltroUser As String
'Dim lsFiltroAge As String
'Dim oConect As COMConecta.DCOMConecta
'
'Set Rs = New ADODB.Recordset
'If psCodUser <> "" Then
'    lsFiltroUser = " AND RIGHT(M.cMovNro,4)  = '" & psCodUser & "' "
'End If
'
'Set oConect = New COMConecta.DCOMConecta
'If oConect.AbreConexion = False Then Exit Function
'
'If psCodCmac = "102" Then
'    'Solo para Lima
'    ssql = "SELECT cGrupoNombre, SUM(nTotalMov) AS nTotalMov, " _
'        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
'        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END ) AS Cheque, " _
'        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, G.cGrupoCod " _
'        & "FROM GrupoOpe G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
'        & "JOIN (SELECT M.cOpeCod, O.cOpeDesc, COUNT(*) AS nTotalMov, " _
'        & "  CASE WHEN MD.nDocTpo IS NULL THEN SUM(MC.nMonto) ELSE 0 END AS Efectivo, " _
'        & "  CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN  SUM(MC.nMonto) ELSE 0 END AS Cheque, " _
'        & "  CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN SUM(MC.nMonto) ELSE 0 END AS OrdenPago " _
'        & "  FROM Mov M JOIN MovCap MC ON MC.nMovNro = M.nMovNro " _
'        & "  LEFT JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
'        & "  JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
'        & "  WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
'        & "  AND SUBSTRING(MC.cCtaCod,9,1)='" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
'        & "  GROUP BY M.cOpeCod, O.cOpeDesc, MD.nDocTpo) As OpeCol  " _
'        & "ON OpeCol.cOpeCod = GO.cOpeCod " _
'        & "WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%' " _
'        & "GROUP BY cGrupoNombre, G.cGrupoCod "
'Else
'    ssql = "SELECT CASE WHEN OpeCol.cCodAge = '" & psCodAge & "' THEN '000' ELSE (OpeCol.cCodAge + '2') END as nAgeOrd, OpeCol.cCodAge, cGrupoNombre, SUM(nTotalMov) AS nTotalMov, " _
'        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
'        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END ) AS Cheque, " _
'        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, G.cGrupoCod " _
'        & "FROM GrupoOpe G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
'        & "JOIN (SELECT SUBSTRING(MC.cCtaCod,4,2) AS cCodAge, M.cOpeCod, O.cOpeDesc, COUNT(*) AS nTotalMov, " _
'        & "  CASE WHEN MD.nDocTpo IS NULL THEN SUM(MC.nMonto) ELSE 0 END AS Efectivo, " _
'        & "  CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN  SUM(MC.nMonto) ELSE 0 END AS Cheque, " _
'        & "  CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN SUM(MC.nMonto) ELSE 0 END AS OrdenPago " _
'        & "  FROM Mov M JOIN MovCap MC ON MC.nMovNro = M.nMovNro " _
'        & "  LEFT JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
'        & "  JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
'        & "  WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
'        & "  AND SUBSTRING(MC.cCtaCod,9,1)='" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
'        & "  GROUP BY M.cOpeCod, O.cOpeDesc, MD.nDocTpo, SUBSTRING(MC.cCtaCod,4,2) ) As OpeCol  " _
'        & "ON OpeCol.cOpeCod = GO.cOpeCod " _
'        & "WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%' " _
'        & "GROUP BY OpeCol.cCodAge, cGrupoNombre, G.cGrupoCod "
'    ssql = ssql + " UNION " _
'        & "(SELECT DISTINCT CASE WHEN cAgeCod = '" & psCodAge & "' THEN '000' ELSE  (cAgeCod + '1') END as nAgeOrd, cAgeCod, cAgeDescripcion,0,0,0,0,'' " _
'        & "FROM Agencias AG JOIN " _
'        & "(SELECT SUBSTRING(MC.cCtaCod,4,2) As cCodAge " _
'        & " FROM MovCap MC JOIN Mov M ON M.nMovNro = MC.nMovNro " _
'        & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
'        & " AND SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
'        & ") Mov ON Mov.cCodAge = Ag.cAgeCod) " _
'        & " ORDER BY nAgeOrd "
'End If
'
'
'Set Rs = oConect.CargaRecordSet(ssql)
'Set GetOpeIngEgreCaptaciones = Rs
'oConect.CierraConexion: Set oConect = Nothing
'End Function
Private Sub OpeDiaCreaTemporal(ByVal psFecha As String, ByVal psCodUsu As String, ByRef conex As COMConecta.DCOMConecta)
Dim sql As String

sql = " if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[##Mov_" & psCodUsu & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" _
& " drop table [dbo].[##Mov_" & psCodUsu & "]" _
& " "
'oCon.Ejecutar Sql
conex.Ejecutar sql

sql = " CREATE TABLE [dbo].[##Mov_" & psCodUsu & "] (" _
& " [nMovNro] [int] NOT NULL ," _
& " [cMovNro] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," _
& " [cOpeCod] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," _
& " [cMovDesc] [varchar] (302) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," _
& " [nMovEstado] [int] NOT NULL ," _
& " [nMovFlag] [int] NULL" _
& " ) ON [PRIMARY]" _
& " "
'oCon.Ejecutar Sql
conex.Ejecutar sql

sql = " ALTER TABLE [dbo].[##Mov_" & psCodUsu & "] WITH NOCHECK ADD" _
& " CONSTRAINT [PK_Mov_" & psCodUsu & "] PRIMARY KEY CLUSTERED" _
& " (" _
& " [nMovNro]" _
& " ) WITH FILLFACTOR = 90 ON [PRIMARY]" _
& " "
'oCon.Ejecutar Sql
conex.Ejecutar sql

sql = " Insert Into [dbo].[##Mov_" & psCodUsu & "] (nMovNro, cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag)" _
& " Select nMovNro, cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag From Mov Where cMovNro Like '" & psFecha & "%'"
'oCon.Ejecutar Sql
conex.Ejecutar sql

'oCon.CierraConexion

End Sub
Private Sub OpeDiaEliminaTemporal(psCodUsu As String, conex As COMConecta.DCOMConecta)
Dim sql As String
'Dim oCon As COMConecta.DCOMConecta
'Set oCon = New COMConecta.DCOMConecta

'oCon.AbreConexion

sql = " drop table [dbo].[##Mov_" & psCodUsu & "]"
'oCon.Ejecutar Sql
conex.Ejecutar sql
'
'oCon.CierraConexion
End Sub

Public Function GetOpeIngEgreCaptaciones(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                                        ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String, _
                                        Optional ByVal psCodCmac As String) As ADODB.Recordset
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim lsFiltroUser As String
    Dim lsFiltroAge As String
    Dim oConect As COMConecta.DCOMConecta
    
    Dim lsTablaDiaria As String
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    
    If pdFecha = pdFechaDia Then
        lsTablaDiaria = " MovDiario "
    Else
        Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
        lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
    End If
    
    Set rs = New ADODB.Recordset
    If psCodUser <> "" Then
        lsFiltroUser = " AND RIGHT(M.cMovNro,4)  = '" & psCodUser & "' "
    End If

   

    If psCodCmac = "102" Then
        'Solo para Lima
        sSQL = "SELECT cGrupoNombre, SUM(nTotalMov) AS nTotalMov, " _
            & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
            & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END ) AS Cheque , " _
            & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, G.cGrupoCod " _
            & "FROM GrupoOpe G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
            & "JOIN (SELECT Mc.cOpeCod, O.cOpeDesc, COUNT(nmovnro) AS nTotalMov, " _
            & "  CASE WHEN MD.nDocTpo IS NULL THEN SUM(MC.nMonto) ELSE 0 END AS Efectivo , " _
            & "  CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN  SUM(MC.nMonto) ELSE 0 END AS Cheque, " _
            & "  CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN SUM(MC.nMonto) ELSE 0 END AS OrdenPago " _
            & "  FROM Mov M JOIN MovCap MC ON MC.nMovNro = M.nMovNro  " _
            & "  JOIN (SELECT NMOVNRO,COPECOD,CCTACOD,NCONCEPTOCOD FROM MOVCAPDET WHERE NCONCEPTOCOD NOT IN (21,3)) MCD ON   MCD.NMOVNRO=MC.NMOVNRO AND MCD.CCTACOD=MC.CCTACOD AND MCD.COPECOD=MC.COPECOD " _
            & "  LEFT JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
            & "  JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
            & "  WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
            & "  AND SUBSTRING(MC.cCtaCod,9,1)='" & pnMoneda & "' And M.nMovFlag =" & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
            & "  GROUP BY MC.cOpeCod, O.cOpeDesc, MD.nDocTpo) As OpeCol  " _
            & "ON OpeCol.cOpeCod = GO.cOpeCod " _
            & "WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%' " _
            & "GROUP BY cGrupoNombre, G.cGrupoCod "
    Else
        sSQL = " SELECT CASE WHEN OpeCol.cCodAge = '" & psCodAge & "' THEN '000' ELSE (OpeCol.cCodAge + '2') END as nAgeOrd, OpeCol.cCodAge, cGrupoNombre, SUM(nTotalMov) AS nTotalMov, " _
            & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
            & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END ) AS Cheque , " _
            & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, G.cGrupoCod " _
            & "FROM GrupoOpe G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
            & "JOIN (SELECT SUBSTRING(MC.cCtaCod,4,2) AS cCodAge, Mc.cOpeCod, O.cOpeDesc, COUNT(DISTINCT m.nmovnro) AS nTotalMov, " _
            & " CASE WHEN (MD.nDocTpo IS NULL OR MC.COPECOD LIKE '99%') THEN SUM(MC.nMonto) ELSE 0 END AS Efectivo , " _
            & " CASE WHEN (MD.nDocTpo = " & TpoDocCheque & " And MC.cOpeCod NOT LIKE '99%') THEN  SUM(MC.nMonto) ELSE 0 END AS Cheque, " _
            & " CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " AND NOT MC.COPECOD LIKE '99%' THEN SUM(MC.nMonto) ELSE 0 END AS OrdenPago " _
            & " FROM " & lsTablaDiaria & " M JOIN MovCap MC ON MC.nMovNro = M.nMovNro JOIN " _
            & " (Select nMovNro, cOpeCod, cCtaCod From MovCapDet Where nConceptoCod NOT IN (21) Group by nMovNro, cOpeCod, cCtaCod) MCD ON MC.nMovNro = MCD.nMovNro And MC.cCtaCod = MCD.cCtaCod And MC.cOpeCod = MCD.cOpeCod " _
            & " LEFT JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
            & " JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
            & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
            & " AND SUBSTRING(MC.cCtaCod,9,1)='" & pnMoneda & "' And M.nMovFlag =" & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
            & " GROUP BY Mc.cOpeCod, O.cOpeDesc, MD.nDocTpo, SUBSTRING(MC.cCtaCod,4,2) "
            
         sSQL = sSQL & "UNION" _
            & " SELECT SUBSTRING(M.cMovNro,18,2) AS cCodAge, M.cOpeCod, O.cOpeDesc, COUNT(m.nmovnro) AS nTotalMov," _
            & " CASE WHEN MD.nDocTpo IS NULL THEN SUM(MV.nMovImporte) ELSE 0 END AS Efectivo , " _
            & " CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN SUM(MV.nMovImporte) ELSE 0 END AS Cheque," _
            & " CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN SUM(MV.nMovImporte) ELSE 0 END AS OrdenPago " _
            & " FROM " & lsTablaDiaria & " M " _
            & " JOIN MovOpevarias MV ON MV.nMovNro = M.nMovNro" _
            & " JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ")" _
            & " JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod" _
            & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "'" & lsFiltroUser & " " _
            & " AND MV.nMoneda = '" & pnMoneda & "' And M.nMovFlag = 0 AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
            & " GROUP BY M.cOpeCod, O.cOpeDesc, MD.nDocTpo, SUBSTRING(M.cMovNro,18,2)"
            
        sSQL = sSQL & ") As OpeCol  " _
            & "ON OpeCol.cOpeCod = GO.cOpeCod " _
            & "WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%' " _
            & "GROUP BY OpeCol.cCodAge, cGrupoNombre, G.cGrupoCod "
        sSQL = sSQL + " UNION " _
            & "(SELECT DISTINCT CASE WHEN cAgeCod = '" & psCodAge & "' THEN '000' ELSE  (cAgeCod + '1') END as nAgeOrd, cAgeCod, cAgeDescripcion,0,0,0,0,'' " _
            & "FROM Agencias AG JOIN " _
            & "(SELECT SUBSTRING(MC.cCtaCod,4,2) As cCodAge " _
            & " FROM MovCap MC JOIN " & lsTablaDiaria & " M ON M.nMovNro = MC.nMovNro " _
            & "  JOIN (SELECT NMOVNRO,COPECOD,CCTACOD,NCONCEPTOCOD FROM MOVCAPDET WHERE NCONCEPTOCOD<>21) MCD ON   MCD.NMOVNRO=MC.NMOVNRO AND MCD.CCTACOD=MC.CCTACOD AND MCD.COPECOD=MC.COPECOD " _
            & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
            & " AND SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
            & ") Mov ON Mov.cCodAge = Ag.cAgeCod) " _
            & " ORDER BY nAgeOrd "
    End If
    
   Set rs = oConect.CargaRecordSet(sSQL)
    Set GetOpeIngEgreCaptaciones = rs
    If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
    End If
    oConect.CierraConexion: Set oConect = Nothing
    End Function


Public Function GetOpeIngEgreColocaciones(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
            ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String, _
            Optional ByVal psCodCmac As String) As ADODB.Recordset

Dim sSQL As String
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim oConect As COMConecta.DCOMConecta
Dim lsTablaDiaria As String

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

If pdFecha = pdFechaDia Then
    lsTablaDiaria = " MovDiario "
Else
    'psCodUser = IIf(psCodUser = "", "9999", psCodUser)
    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
End If


Set rs = New ADODB.Recordset

If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.cMovNro,4)  = '" & psCodUser & "' "
End If



If psCodCmac = "102" Then
    'Solo para Lima
    sSQL = "SELECT cGrupoNombre, SUM(nTotalMov) AS nTotalMov, " _
        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END ) AS Cheque, " _
        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, G.cGrupoCod " _
        & "FROM GrupoOpe G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
        & "JOIN (SELECT M.cOpeCod, O.cOpeDesc, COUNT(*) AS nTotalMov, " _
        & "  CASE WHEN MD.nDocTpo IS NULL THEN SUM(MC.nMonto) ELSE 0 END AS Efectivo, " _
        & "  CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN  SUM(MC.nMonto) ELSE 0 END AS Cheque, " _
        & "  CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN SUM(MC.nMonto) ELSE 0 END AS OrdenPago " _
        & "  FROM Mov M JOIN MovCol MC ON MC.nMovNro = M.nMovNro " _
        & "  LEFT JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
        & "  JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
        & "  WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
        & "  AND SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
        & "  GROUP BY M.cOpeCod, O.cOpeDesc, MD.nDocTpo) As OpeCol  " _
        & "ON OpeCol.cOpeCod = GO.cOpeCod " _
        & "WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%' " _
        & "GROUP BY cGrupoNombre, G.cGrupoCod "
Else
    sSQL = "SELECT CASE WHEN OpeCol.cCodAge = '" & psCodAge & "' THEN '000' ELSE (OpeCol.cCodAge + '2') END as nAgeOrd, OpeCol.cCodAge, cGrupoNombre, SUM(nTotalMov) AS nTotalMov, " _
        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END ) AS Cheque, " _
        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, G.cGrupoCod " _
        & "FROM GrupoOpe G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
        & "JOIN (SELECT SUBSTRING(MC.cCtaCod,4,2) AS cCodAge, MC.cOpeCod, O.cOpeDesc, COUNT(M.NMOVNRO) AS nTotalMov, " _
        & "  CASE WHEN MD.nDocTpo IS NULL THEN SUM(MC.nMonto) ELSE 0 END AS Efectivo, " _
        & "  CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN  SUM(MC.nMonto) ELSE 0 END AS Cheque, " _
        & "  CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN SUM(MC.nMonto) ELSE 0 END AS OrdenPago " _
        & "  FROM " & lsTablaDiaria & " M JOIN MovCol MC ON MC.nMovNro = M.nMovNro " _
        & "  LEFT JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
        & "  JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
        & "  WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
        & "  AND SUBSTRING(MC.cCtaCod,9,1)='" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
        & "  GROUP BY MC.cOpeCod, O.cOpeDesc, MD.nDocTpo, SUBSTRING(MC.cCtaCod,4,2) ) As OpeCol  " _
        & "ON OpeCol.cOpeCod = GO.cOpeCod " _
        & "WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%' " _
        & "GROUP BY OpeCol.cCodAge, cGrupoNombre, G.cGrupoCod "
    sSQL = sSQL + " UNION " _
        & "(SELECT DISTINCT CASE WHEN cAgeCod = '" & psCodAge & "' THEN '000' ELSE  (cAgeCod + '1') END as nAgeOrd, cAgeCod, cAgeDescripcion,0,0,0,0,'' " _
        & "FROM Agencias AG JOIN " _
        & "(SELECT SUBSTRING(MC.cCtaCod,4,2) As cCodAge " _
        & " FROM MovCol MC JOIN " & lsTablaDiaria & " M ON M.nMovNro = MC.nMovNro " _
        & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser & lsFiltroAge _
        & " AND SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And M.nMovFlag =" & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
        & ") Mov ON Mov.cCodAge = Ag.cAgeCod) " _
        & " ORDER BY nAgeOrd "
End If

Set rs = oConect.CargaRecordSet(sSQL)
Set GetOpeIngEgreColocaciones = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function

'Public Function GetOpeIngEgreOtrasOpe(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
'                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda) As ADODB.Recordset
'Dim sql As String
'Dim Rs As ADODB.Recordset
'Dim lsFiltroUser As String
'Dim lsFiltroAge As String
'Dim oConect As COMConecta.DCOMConecta
'
'Set Rs = New ADODB.Recordset
'If psCodAge <> "" Then
'    lsFiltroAge = " AND  SUBSTRING(M.CMOVNRO,18,2) = '" & psCodAge & "'"
'End If
'If psCodUser <> "" Then
'   If gsCodCMAC = "102" Then 'CASL 11/11/2003 Solo para Lima
'      If psCodUser = "BOVE" Then
'         lsFiltroUser = " AND (RIGHT(M.CMOVNRO,4) = '" & psCodUser & "' "
'         lsFiltroUser = lsFiltroUser & " OR M.cOpeCod = '" & gOtrOpeDepCtaBcoEfec & "') "
'      Else
'         lsFiltroUser = " AND RIGHT(M.CMOVNRO,4) = '" & psCodUser & "' "
'         lsFiltroUser = lsFiltroUser & " AND not M.cOpeCod = '" & gOtrOpeDepCtaBcoEfec & "' "
'      End If
'   Else
'         lsFiltroUser = " AND (RIGHT(M.CMOVNRO,4) = '" & psCodUser & "' "
'   End If
'End If
'Set oConect = New COMConecta.DCOMConecta
'If oConect.AbreConexion = False Then Exit Function
'
''--Otras Operaciones
'sql = "SELECT cGrupoNombre, cCodAge = '" & psCodAge & "', SUM(nTotalMov) AS nTotalMov, " _
'    & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN EFECTIVO ELSE EFECTIVO*-1 END) AS EFECTIVO, " _
'    & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN CHEQUE ELSE CHEQUE*-1 END) AS CHEQUE, " _
'    & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN ORDENPAGO ELSE ORDENPAGO*-1 END) AS ORDENPAGO, GO.cGrupoCod " _
'    & "FROM GRUPOOPE  AS G JOIN GRUPOSOPE GO ON GO.cGrupoCod=G.cGrupoCod " _
'    & "INNER JOIN (SELECT M.cOpeCod, O.COPEDESC, COUNT(*) AS nTotalMov, " _
'    & "  CASE WHEN MD.nDocTpo IS NULL THEN  SUM(MV.nMovImporte) ELSE 0 END AS EFECTIVO, " _
'    & "  CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN  SUM(MV.nMovImporte) ELSE 0 END AS CHEQUE, " _
'    & "  CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN  SUM(MV.nMovImporte) ELSE 0 END AS ORDENPAGO " _
'    & "  FROM MOV M JOIN MOVOPEVARIAS MV ON MV.nMovNro = M.nMovNro " _
'    & "  LEFT JOIN MOVDOC MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
'    & "  JOIN OPETPO O ON O.cOpeCod = M.cOpeCod " _
'    & "  WHERE   SUBSTRING(M.CMOVNRO,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser _
'    & "  AND M.NMOVFLAG =" & gMovFlagVigente & " AND MV.nMoneda =" & pnMoneda & " " _
'    & "  GROUP BY  M.cOpeCod, O.COPEDESC, MD.nDocTpo ) AS OPESERV " _
'    & "  ON OPESERV.cOpeCod = GO.cOpeCod " _
'    & " WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%'  " _
'    & " GROUP BY cGrupoNombre, GO.cGrupoCod "
'
'Set Rs = oConect.CargaRecordSet(sql)
'Set GetOpeIngEgreOtrasOpe = Rs
'oConect.CierraConexion: Set oConect = Nothing
'End Function

Public Function GetOpeIngEgreOtrasOpe(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim lsFiltroUser2 As String
Dim lsFiltroAge2 As String
Dim oConect As COMConecta.DCOMConecta
Dim lsTablaDiaria As String

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function


If pdFecha = pdFechaDia Then
    lsTablaDiaria = " MovDiario "
Else
    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
End If

Set rs = New ADODB.Recordset
If psCodAge <> "" Then
    lsFiltroAge = " AND  SUBSTRING(M.CMOVNRO,18,2) = '" & psCodAge & "'"
    lsFiltroAge2 = " AND  SUBSTRING(CMOVNRO,18,2) = '" & psCodAge & "'"
End If
If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.CMOVNRO,4) = '" & psCodUser & "' "
    lsFiltroUser2 = " AND RIGHT(CMOVNRO,4) = '" & psCodUser & "' "
End If

'--Otras Operaciones
sql = "SELECT cGrupoNombre, cCodAge = '" & psCodAge & "', SUM(nTotalMov) AS nTotalMov, " _
    & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN EFECTIVO ELSE EFECTIVO*-1 END) AS EFECTIVO, " _
    & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN CHEQUE ELSE CHEQUE*-1 END) AS CHEQUE, " _
    & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN ORDENPAGO ELSE ORDENPAGO*-1 END) AS ORDENPAGO, GO.cGrupoCod " _
    & "FROM GRUPOOPE  AS G JOIN GRUPOSOPE GO ON GO.cGrupoCod=G.cGrupoCod " _
    & "INNER JOIN (SELECT M.cOpeCod, O.COPEDESC, COUNT(*) AS nTotalMov, " _
    & "  CASE WHEN MD.nDocTpo IS NULL THEN  SUM(MV.nMovImporte) ELSE 0 END AS EFECTIVO, " _
    & "  CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN  SUM(MV.nMovImporte) ELSE 0 END AS CHEQUE, " _
    & "  CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN  SUM(MV.nMovImporte) ELSE 0 END AS ORDENPAGO " _
    & "  FROM " & lsTablaDiaria & " M JOIN MOVOPEVARIAS MV ON MV.nMovNro = M.nMovNro " _
    & "  LEFT JOIN MOVDOC MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
    & "  JOIN OPETPO O ON O.cOpeCod = M.cOpeCod " _
    & "  WHERE   SUBSTRING(M.CMOVNRO,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser _
    & "  AND M.NMOVFLAG =" & gMovFlagVigente & " AND MV.nMoneda =" & pnMoneda & " " _
    & "  GROUP BY  M.cOpeCod, O.COPEDESC, MD.nDocTpo ) AS OPESERV " _
    & "  ON OPESERV.cOpeCod = GO.cOpeCod " _
    & " WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%'  " _
    & " GROUP BY cGrupoNombre, GO.cGrupoCod "
    
'--Unimos otras operaciones con ITF
'sql = sql & " UNION "
'sql = sql & "Select 'Recaudación I.T.F.' As cGrupoNombre , A.cAgeCod as cCodAge, SUM(nTotalMov) nTotalMov, sum(ITF) EFECTIVO, 0 as CHEQUE, 0 as ORDENPAGO, '0' cGrupoCod from " _
'    & " (select  SubString(cMovNro,18,2) Agencia, " _
'    & "         Count(nMovITF) nTotalMov, " _
'    & " sum(Round(nMontoITF,2)) ITF " _
'    & " from MovITF where cCtaCod IS NOT NULL " _
'    & " And cOpeCod not in ('200302','200304','200305','200306','200307','200308','100205','100305','100405','260101','260103','260104') " _
'    & " And nExtornoITF=0 " _
'    & " And SUBSTRING(CMOVNRO,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " _
'    & lsFiltroAge2 & lsFiltroUser2 _
'    & " And SubString(cCtaCod,9,1) = '" & pnMoneda & "'" _
'    & " Group by SubString(cMovNro,18,2),  Right(cMovNro,4) " _
'    & " ) GG " _
'    & " inner join Agencias A on A.cAgeCod=GG.Agencia " _
'    & " group by A.cAgeCod " _
'    & " order by cCodAge "

Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreOtrasOpe = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetOpeIngEgreServicios(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim oConect As COMConecta.DCOMConecta
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function


Dim lsTablaDiaria As String

If pdFecha = pdFechaDia Then
    lsTablaDiaria = " MovDiario "
Else
    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
End If

Set rs = New ADODB.Recordset
If psCodAge <> "" Then
    lsFiltroAge = " AND  SUBSTRING(M.CMOVNRO,18,2) = '" & psCodAge & "'"
End If
If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.CMOVNRO,4) = '" & psCodUser & "' "
End If

'--Servicios
sql = "SELECT cGrupoNombre, cCodAge = '" & psCodAge & "', SUM(nTotalMov) AS nTotalMov, " _
    & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
    & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END) AS Cheque, " _
    & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, GO.cGrupoCod " _
    & "FROM GRUPOOPE AS G JOIN GRUPOSOPE GO ON GO.cGrupoCod = G.cGrupoCod " _
    & "INNER JOIN (SELECT M.cOpeCod, O.COPEDESC, COUNT(*) AS nTotalMov, SUM(MV.nMonto) Efectivo, " _
    & "  Cheque = 0, OrdenPago = 0 FROM " & lsTablaDiaria & " M JOIN MovServicios MV ON MV.nMovNro = M.nMovNro " _
    & "  JOIN OPETPO O ON O.cOpeCod = M.cOpeCod " _
    & "  WHERE SUBSTRING(M.CMOVNRO,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser _
    & "  AND M.NMOVFLAG = " & gMovFlagVigente & " And MV.nMoneda = " & pnMoneda & " " _
    & "  GROUP BY  M.cOpeCod, O.COPEDESC) AS OPESERV " _
    & "  ON OPESERV.cOpeCod = GO.cOpeCod " _
    & " WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%'  " _
    & " GROUP BY cGrupoNombre, GO.cGrupoCod "

Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreServicios = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetOpeIngEgreSobranteFaltante(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String, _
                Optional ByVal psUsuarioBOVEDA As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim oConect As COMConecta.DCOMConecta
Dim lsTablaDiaria As String

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function



If pdFecha = pdFechaDia Then
    lsTablaDiaria = " MovDiario "
Else
    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
End If

Set rs = New ADODB.Recordset
If psCodAge <> "" Then
    lsFiltroAge = " AND  SUBSTRING(M.CMOVNRO,18,2) = '" & psCodAge & "'"
End If

''If psCodUser <> "" And psCodUser <> psUsuarioBOVEDA Then
 '   lsFiltroUser = " AND RIGHT(M.CMOVNRO,4) = '" & psCodUser & "' "
'End If

If psCodUser <> "" And psCodUser <> psUsuarioBOVEDA Then
    lsFiltroUser = " AND RIGHT(M.CMOVNRO,4) = '" & psCodUser & "' and not M.cOpeCod IN ('" & COMDConstSistema.gOpeBoveAgeRegFaltante & "','" & COMDConstSistema.gOpeBoveAgeRegSobrante & "')"
ElseIf psCodUser = psUsuarioBOVEDA Then
    lsFiltroUser = " AND M.cOpeCod IN ('" & COMDConstSistema.gOpeBoveAgeRegFaltante & "','" & COMDConstSistema.gOpeBoveAgeRegSobrante & "') "
End If

'--Otras Operaciones
sql = "SELECT cGrupoNombre, SUM(nTotalMov) AS nTotalMov, cCodAge = '" & psCodAge & "', " _
    & "SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
    & "Cheque = 0, OrdenPago = 0, GO.cGrupoCod " _
    & " FROM GRUPOOPE AS G JOIN GRUPOSOPE GO ON GO.cGrupoCod = G.cGrupoCod " _
    & " JOIN (SELECT  M.cOpeCod, O.COPEDESC, COUNT(*) AS nTotalMov, SUM(Abs(MV.nMovImporte)) As Efectivo " _
    & " FROM  " & lsTablaDiaria & " M JOIN MOVOPEVARIAS MV ON MV.NMOVNRO = M.NMOVNRO " _
    & " JOIN OPETPO O ON O.cOpeCod = M.cOpeCod " _
    & " WHERE SUBSTRING(M.CMOVNRO,1,8) ='" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser _
    & " AND M.nMovFlag = " & gMovFlagVigente & " AND MV.nMoneda = " & pnMoneda & " AND ABS(MV.nMovImporte) > 0 " _
    & " GROUP BY  M.cOpeCod, O.cOpeDesc ) AS OPESERV " _
    & " ON OPESERV.cOpeCod = GO.cOpeCod " _
    & " WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%'  " _
    & " GROUP BY cGrupoNombre, GO.cGrupoCod "

Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreSobranteFaltante = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetOpeIngEgreCompraVenta(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                        ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String) As ADODB.Recordset

Dim sql As String
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim oConect As COMConecta.DCOMConecta
Dim lsTablaDiaria As String

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function


If pdFecha = pdFechaDia Then
    lsTablaDiaria = " MovDiario "
Else
    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
End If

Set rs = New ADODB.Recordset
If psCodAge <> "" Then
    lsFiltroAge = " AND  SUBSTRING(M.CMOVNRO,18,2) ='" & psCodAge & "'"
End If
If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.CMOVNRO,4) = '" & psCodUser & "' "
End If


'--COMPRA VENTA ME  '--si el reporte es en dolares se debe mostrar solo el campo EFECDOL
'--por ahora sólo mustro la parte soles puesto que sólo estoy mostrando la consulta en soles
'--PARA PODER REALIZAR EL UNION CON LAS DEMAS CONSULTAS
'-- EL SIGNO NO SE BE DIFERENCIAR POR EL GRUPO OPERACIONES PUESTO QUE YA SE DEDUJO A PARTIR
'-- DEL CODIGO DE LA OPERACION MAS NO DEL CAMPO DEL GRUPOOPERACIONES
'--si el reporte es en soles se debe mostrar solo el campo EFECSOLES"
sql = "SELECT cGrupoNombre, SUM(nTotalMov) nTotalMov, cCodAge = '" & psCodAge & "', " _
    & IIf(pnMoneda = gMonedaNacional, " SUM(EfecSoles) ", " SUM(EfecDol) ") & " AS Efectivo, SUM(Cheque) Cheque, " _
    & "SUM(OrdenPago) OrdenPago, GO.cGrupoCod " _
    & "FROM GrupoOpe AS G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
    & "INNER JOIN (SELECT cOpeCod, cOpeDesc, nTotalMov, EfecSoles, EfecDol, 0 AS CHEQUE, 0 AS ORDENPAGO " _
    & "  FROM (SELECT cOpeCod, cOpeDesc, nTotalMov, " _
    & "  CASE WHEN cOpeCod IN ('" & gOpeCajeroMECompra & "','" & gOpeCajeroMECompraEsp & "') THEN EfectivoSoles*-1 ELSE EfectivoSoles END AS EfecSoles, " _
    & "  CASE WHEN cOpeCod IN ('" & gOpeCajeroMECompra & "','" & gOpeCajeroMECompraEsp & "') THEN EfectivoDolares ELSE EfectivoDolares*-1 END AS EfecDol " _
    & "  From (SELECT M.cOpeCod, O.cOpeDesc, COUNT(*) AS nTotalMov, " _
    & "        SUM(ROUND(MC.nMovImporte*TC.nMovTpoCambio,2)) AS  EfectivoSoles, " _
    & "        SUM(MC.nMovImporte) As EfectivoDolares " _
    & "        FROM " & lsTablaDiaria & " M JOIN MOVCOMPRAVENTA MC ON MC.nMovNro = M.nMovNro " _
    & "        JOIN MOVTPOCAMBIO TC ON TC.nMovNro = M.nMovNro " _
    & "        JOIN OPETPO O ON O.cOpeCod = M.cOpeCod " _
    & "        WHERE SUBSTRING(M.CMOVNRO,1,8) ='" & Format$(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser _
    & "        AND M.NMOVFLAG = " & gMovFlagVigente & " " _
    & "     GROUP BY M.cOpeCod, O.cOpeDesc, TC.nMovTpoCambio) AS MTC ) AS TC) AS MTC " _
    & " ON MTC.cOpeCod = GO.cOpeCod Group by G.cGrupoNombre, GO.cGrupoCod"

Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreCompraVenta = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetOpeIngEgreHabDev(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String) As ADODB.Recordset

Dim sSQL As String
Dim rs As ADODB.Recordset
Dim lsFiltroAge As String
Dim oConect As COMConecta.DCOMConecta
Dim lsTablaDiaria As String

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

If pdFecha = pdFechaDia Then
    lsTablaDiaria = " MovDiario "
Else
    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
End If

Set rs = New ADODB.Recordset


'sSql = "SELECT cGrupoNombre, GO.cGrupoCod, SUM(nTotalMov) nTotalMov, cCodAge = '" & psCodAge & "', " _
'    & "SUM(nMontoHabDev) As Efectivo, SUM(Cheque) As Cheque, SUM(OrdenPago) As OrdenPago " _
'    & "FROM GrupoOpe As G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
'    & "JOIN (SELECT cOpeCod, " _
'    & "  CASE WHEN cAgeOrig = '" & psCodAge & "' THEN  Efectivo*-1 ELSE Efectivo END AS nMontoHabDev, " _
'    & "  0 AS Cheque, 0 AS OrdenPago, nTotalMov " _
'    & "  FROM (SELECT M.cOpeCod, COUNT(*) AS nTotalMov, SUM(MH.nMovImporte) AS Efectivo, " _
'    & "     ISNULL(MH.cAgeOrig,'') AS cAgeOrig " _
'    & "     FROM Mov M JOIN MovHabilitacion MH ON MH.nMovNro = M.nMovNro " _
'    & "     JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
'    & " WHERE  " _
'    & "     M.nMovFlag = " & gMovFlagVigente & " And MH.nMoneda = " & pnMoneda & " " _
'    & "     AND (MH.cAgeOrig = '" & psCodAge & "' OR MH.cAgeDest = '" & psCodAge & "')" _
'    & "     AND EXISTS (SELECT M.nMovNro FROM MOV M1 JOIN MOVREF MR ON MR.nMovNro = M1.nMovNro " _
'    & "     WHERE SUBSTRING(M1.cMovNro,1,8) = '" & Format$(pdFecha, "yyyymmdd") & "' AND MR.nMovNroREF = M.nMovNro AND M1.nMovFlag = " & gMovFlagVigente & ") " _
'    & "     GROUP BY M.cOpeCod, MH.cAgeOrig) AS HabDev) As OpeHD " _
'    & "ON OpeHD.cOpeCod = GO.cOpeCod Where G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%' " _
'    & "GROUP BY cGrupoNombre, GO.cGrupoCod  "

sSQL = "SELECT cGrupoNombre, GO.cGrupoCod, SUM(nTotalMov) nTotalMov, cCodAge = '" & psCodAge & "', " _
    & "SUM(nMontoHabDev) As Efectivo, SUM(Cheque) As Cheque, SUM(OrdenPago) As OrdenPago " _
    & "FROM GrupoOpe As G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
    & "JOIN (SELECT cOpeCod, " _
    & "  CASE WHEN cAgeOrig = '" & psCodAge & "' THEN  Efectivo*-1 ELSE Efectivo END AS nMontoHabDev, " _
    & "  0 AS Cheque, 0 AS OrdenPago, nTotalMov " _
    & "  FROM (SELECT M.cOpeCod, COUNT(*) AS nTotalMov, SUM(MH.nMovImporte) AS Efectivo, " _
    & "     ISNULL(MH.cAgeOrig,'') AS cAgeOrig " _
    & "     FROM " & lsTablaDiaria & " M JOIN MovHabilitacion MH ON MH.nMovNro = M.nMovNro " _
    & "     JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
    & " WHERE  " _
    & "     M.nMovFlag = " & gMovFlagVigente & " And MH.nMoneda = " & pnMoneda & " " _
    & "     AND (MH.cAgeOrig = '" & psCodAge & "' )" _
    & "     AND SUBSTRING(M.cMovNro,1,8) = '" & Format$(pdFecha, "yyyymmdd") & "' " _
    & "     GROUP BY M.cOpeCod, MH.cAgeOrig) AS HabDev"
sSQL = sSQL & " UNION SELECT cOpeCod, " _
    & "  CASE WHEN cAgeOrig = '" & psCodAge & "' THEN  Efectivo*-1 ELSE Efectivo END AS nMontoHabDev, " _
    & "  0 AS Cheque, 0 AS OrdenPago, nTotalMov " _
    & "  FROM (SELECT M.cOpeCod, COUNT(*) AS nTotalMov, SUM(MH.nMovImporte) AS Efectivo, " _
    & "     ISNULL(MH.cAgeOrig,'') AS cAgeOrig " _
    & "     FROM " & lsTablaDiaria & " M JOIN MovHabilitacion MH ON MH.nMovNro = M.nMovNro " _
    & "     JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
    & " WHERE  " _
    & "     M.nMovFlag = " & gMovFlagVigente & " And MH.nMoneda = " & pnMoneda & " " _
    & "     AND (MH.cAgeDest = '" & psCodAge & "')" _
    & "     AND EXISTS (SELECT M.nMovNro FROM " & lsTablaDiaria & " M1 JOIN MOVREF MR ON MR.nMovNro = M1.nMovNro " _
    & "     WHERE SUBSTRING(M1.cMovNro,1,8) = '" & Format$(pdFecha, "yyyymmdd") & "' AND MR.nMovNroREF = M.nMovNro AND M1.nMovFlag = " & gMovFlagVigente & ") " _
    & "     GROUP BY M.cOpeCod, MH.cAgeOrig) AS HabDev ) As OpeHD " _
    & "ON OpeHD.cOpeCod = GO.cOpeCod Where G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%' " _
    & "GROUP BY cGrupoNombre, GO.cGrupoCod  "

Set rs = oConect.CargaRecordSet(sSQL)
Set GetOpeIngEgreHabDev = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetOpeIngEgreHabDevCajero(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String) As ADODB.Recordset

Dim sSQL As String
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim oConect As COMConecta.DCOMConecta
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function


Dim lsTablaDiaria As String

If pdFecha = pdFechaDia Then
    lsTablaDiaria = " MovDiario "
Else
    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
End If

Set rs = New ADODB.Recordset


sSQL = "SELECT cGrupoNombre, GO.cGrupoCod, Sum(nTotalMov) nTotalMov, cCodAge = '" & psCodAge & "', " _
    & "Efectivo = CASE WHEN G.cIngEgr = 'I' THEN SUM(nMontoHabDev) ELSE SUM(nMontoHabDev)*-1 END, " _
    & "SUM(Cheque) As Cheque, SUM(OrdenPago) As OrdenPago FROM GrupoOpe G JOIN GruposOpe GO " _
    & "ON GO.cGrupoCod = G.cGrupoCod JOIN " _
    & "(SELECT cOpeCod, nTotalMov, Efectivo AS nMontoHabDev, 0 AS Cheque , 0 AS OrdenPago FROM " _
    & "(SELECT M1.cOpeCod, COUNT(*) AS nTotalMov, SUM(MH.nMovImporte) AS Efectivo " _
    & "FROM OpeTpo O JOIN " & lsTablaDiaria & " M JOIN MovHabDevCajero MH ON MH.nMovNro = M.nMovNro ON O.cOpeCod = M.cOpeCod " _
    & "JOIN MovRef MR JOIN  " & lsTablaDiaria & "  M1 ON MR.nMovNro = M1.nMovNro ON MR.nMovNroRef = M.nMovNro " _
    & "WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format$(pdFecha, "yyyymmdd") & "' AND MH.cUsuDest = '" & psCodUser & "' AND " _
    & "M.nMovFlag = 0 AND MH.cAgeCod = '" & psCodAge & "' And MH.nMoneda = " & pnMoneda & " And M1.nMovFlag = " & gMovFlagVigente & " " _
    & "Group by M1.cOpeCod UNION " _
    & "SELECT M.cOpeCod, COUNT(*) AS nTotalMov, SUM(MH.nMovImporte) AS Efectivo " _
    & "FROM OpeTpo O JOIN " & lsTablaDiaria & " M JOIN MovHabDevCajero MH ON MH.nMovNro = M.nMovNro ON O.cOpeCod = M.cOpeCod " _
    & "JOIN MovRef MR JOIN " & lsTablaDiaria & " M1 ON MR.nMovNro = M1.nMovNro ON MR.nMovNroRef = M.nMovNro " _
    & "WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format$(pdFecha, "yyyymmdd") & "' AND MH.cUsuOrig = '" & psCodUser & "' AND " _
    & "M.nMovFlag = 0 AND MH.cAgeCod = '" & psCodAge & "' And MH.nMoneda = " & pnMoneda & " And M1.nMovFlag = " & gMovFlagVigente & " " _
    & "Group by M.cOpeCod) HabDev) OpeHD ON OpeHD.cOpeCod = GO.cOpeCod Where " _
    & "G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%' GROUP BY cGrupoNombre, GO.cGrupoCod, G.cIngEgr"

Set rs = oConect.CargaRecordSet(sSQL)
Set GetOpeIngEgreHabDevCajero = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If

oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetOpeIngEgreHabDevBillCajero(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String) As ADODB.Recordset

Dim sSQL As String
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim oConect As COMConecta.DCOMConecta
Dim oVar As New COMFunciones.FCOMVarPublicas
Dim lsTablaDiaria As String

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function


If pdFecha = pdFechaDia Then
    lsTablaDiaria = " MovDiario "
Else
    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
End If

Set rs = New ADODB.Recordset


sSQL = "SELECT cGrupoNombre, GO.cGrupoCod, Sum(nTotalMov) nTotalMov, cCodAge = '" & psCodAge & "', " _
    & "Efectivo = CASE WHEN G.cIngEgr = 'I' THEN SUM(nMontoHabDev) ELSE SUM(nMontoHabDev)*-1 END, " _
    & "SUM(Cheque) As Cheque, SUM(OrdenPago) As OrdenPago FROM GrupoOpe G JOIN GruposOpe GO " _
    & "ON GO.cGrupoCod = G.cGrupoCod JOIN " _
    & "(SELECT cOpeCod, nTotalMov, Efectivo AS nMontoHabDev, 0 AS Cheque , 0 AS OrdenPago FROM " _
    & "(SELECT M.cOpeCod, COUNT(*) AS nTotalMov, SUM(MH.nMovImporte) AS Efectivo " _
    & "FROM OpeTpo O JOIN " & lsTablaDiaria & " M JOIN MovHabDevCajero MH ON MH.nMovNro = M.nMovNro ON O.cOpeCod = M.cOpeCod " _
    & "WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format$(pdFecha, "yyyymmdd") & "' AND MH.cUsuOrig = '" & psCodUser & "' AND " _
    & "M.nMovFlag = 0 AND MH.cAgeCod = '" & psCodAge & "' And MH.nMoneda = " & pnMoneda & " " _
    & "Group by M.cOpeCod) HabDev) OpeHD ON OpeHD.cOpeCod = GO.cOpeCod Where " _
    & "G.cGrupoCod IN ('" & oVar.gsCajDevBilletaje & "') GROUP BY cGrupoNombre, GO.cGrupoCod, G.cIngEgr"

Set rs = oConect.CargaRecordSet(sSQL)
Set GetOpeIngEgreHabDevBillCajero = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetCompraVentaRepo(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                              ByVal psCodUser As String, ByVal psCodAge As String, Optional lbHistorica As Boolean = False) As ADODB.Recordset
Dim sql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
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
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim lsOpeCod As String

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta

If oConect.AbreConexion = False Then Exit Function
'Select Case psOpeCod
'    Case gOpeHabCajIngEfectRegulaFaltMN
'        lsOpeCod = gOpeHabCajRegSobFaltMN
'    Case gOpeHabCajIngEfectRegulaFaltME
'        lsOpeCod = gOpeHabCajRegSobFaltME
'End Select

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
                ByVal pnImporte As Double, ByVal psReferencia As String, ByVal pnMovNroReg As Long) As Integer
Dim oMOv As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMOv = New COMDMov.DCOMMov

On Error GoTo GrabaRegularizaFaltanteErr
oMOv.BeginTrans
lbTrans = True
GrabaRegularizaFaltante = 1
oMOv.InsertaMov psMovNro, psOpeCod, psMovDesc
lnMovNro = oMOv.GetnMovNro(psMovNro)
oMOv.InsertaMovOpeVarias lnMovNro, "", Trim(psReferencia), pnImporte
oMOv.InsertaMovRef lnMovNro, pnMovNroReg
oMOv.CommitTrans
lbTrans = False
GrabaRegularizaFaltante = 0
Set oMOv = Nothing
Exit Function
GrabaRegularizaFaltanteErr:
    If lbTrans Then
        oMOv.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GrabaRegistroSobranteFaltante(ByVal sMovNro As String, ByVal sOpeCod As String, _
        ByVal sMovDesc As String, ByVal nImporte As Double, ByVal nMoneda As Moneda) As Integer
Dim oMOv As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim nMovNro As Long
Set oMOv = New COMDMov.DCOMMov

On Error GoTo GrabaRegularizaFaltanteErr
oMOv.BeginTrans
lbTrans = True
GrabaRegistroSobranteFaltante = 1
oMOv.InsertaMov sMovNro, sOpeCod, sMovDesc
nMovNro = oMOv.GetnMovNro(sMovNro)
oMOv.InsertaMovOpeVarias nMovNro, "", Trim(sMovDesc), nImporte, nMoneda
oMOv.CommitTrans
lbTrans = False
GrabaRegistroSobranteFaltante = 0
Set oMOv = Nothing
Exit Function
GrabaRegularizaFaltanteErr:
    If lbTrans Then
        oMOv.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ImprimeBilletaje(ByVal rsBillMN As Recordset, ByVal rsBillME As Recordset, _
        ByVal rsMonMN As Recordset, ByVal rsMonME As Recordset, _
        ByVal sNomAge As String, ByVal sUsuario As String, ByVal sNomUsuario As String, _
        ByVal dFecSis As Date) As String

Dim i As Integer, nCarLin As Long
Dim sCad As String
Dim sNumPag As String
Dim sTitRp1 As String, sTitRp2 As String
Dim sTotalMonto As String * 16
Dim nTotalMon As Double, nTotalBill As Double
Dim sMoneda As String, sMonedaTit As String
Dim sSimbolo As String
Dim oImpre As New COMNCaptaGenerales.NCOMCaptaImpresion

sCad = ""

nCarLin = 85

sTitRp1 = "DESCOMPOSICION DE EFECTIVO - " & sUsuario
sTitRp2 = sNomUsuario

sCad = sCad & oImpre.CabeRepoCaptac("", "", nCarLin, "SECCION OPERACIONES", sTitRp1, sTitRp2, sMoneda, "1", Trim(sNomAge), dFecSis) & Chr$(10)

sMoneda = "SOLES"
sMonedaTit = "MONEDA NACIONAL"
sSimbolo = "S/."
sCad = sCad & sMonedaTit & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & ImprimeRsBilletaje(rsBillMN, nTotalBill)
RSet sTotalMonto = Format$(nTotalBill, "#,##0.00")
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & "SUB TOTAL BILLETAJE" & Space(23) & sTotalMonto & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & ImprimeRsBilletaje(rsMonMN, nTotalMon)
RSet sTotalMonto = Format$(nTotalMon, "#,##0.00")
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & "SUB TOTAL MONEDA" & Space(26) & sTotalMonto & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)
RSet sTotalMonto = sSimbolo & " " & Format$(nTotalBill + nTotalMon, "#,##0.00")
sCad = sCad & "TOTAL " & sMonedaTit & Space(21) & sTotalMonto & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & Chr$(10)

sMoneda = "DOLARES"
sMonedaTit = "MONEDA EXTRANJERA"
sSimbolo = "US$"
sCad = sCad & sMonedaTit & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & ImprimeRsBilletaje(rsBillME, nTotalBill)
RSet sTotalMonto = Format$(nTotalBill, "#,##0.00")
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & "SUB TOTAL BILLETAJE" & Space(23) & sTotalMonto & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & ImprimeRsBilletaje(rsMonME, nTotalMon)
RSet sTotalMonto = Format$(nTotalMon, "#,##0.00")
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & "SUB TOTAL MONEDA" & Space(26) & sTotalMonto & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)
RSet sTotalMonto = sSimbolo & " " & Format$(nTotalBill + nTotalMon, "#,##0.00")
sCad = sCad & "TOTAL " & sMonedaTit & Space(21) & sTotalMonto & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)
ImprimeBilletaje = sCad
End Function

Private Function ImprimeRsBilletaje(ByVal rs As Recordset, ByRef nTotalMonto As Double) As String
Dim sCantidad As String * 8
Dim sMonto As String * 16
Dim sDescripcion As String * 30
Dim sCad As String
sCad = ""
nTotalMonto = 0
If Not rs Is Nothing Then
    Do While Not rs.EOF
        RSet sCantidad = Trim(rs("Cantidad"))
        nTotalMonto = nTotalMonto + rs("Monto")
        RSet sMonto = Trim(rs("Monto"))
        sDescripcion = Trim(rs("Descripción"))
        sCad = sCad & sDescripcion & Space(2) & sCantidad & Space(2) & sMonto & Chr$(10)
        rs.MoveNext
    Loop
End If
ImprimeRsBilletaje = sCad
End Function

Public Function YaRealizoDevBilletaje(ByVal sUsuario As String, ByVal dFecha As String, Optional ByVal psCodAge As String = "") As Boolean
Dim sSQL As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

sSQL = "Select M.cOpeCod From Mov M INNER JOIN MovHabDevCajero H ON M.nMovNro = H.nMovNro " _
    & "Where M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And cOpeCod = '" & gOpeHabCajDevBilletaje & "'" _
    & "And H.cUsuOrig = '" & sUsuario & "' And M.nMovFlag IN (" & gMovFlagVigente & ")"

If psCodAge <> "" Then
    sSQL = sSQL & " AND Substring(cMovNro, 18,2) = '" & psCodAge & "'"
End If

Set rs = oConect.CargaRecordSet(sSQL)
If rs.EOF And rs.BOF Then
    YaRealizoDevBilletaje = False
Else
    YaRealizoDevBilletaje = True
End If
rs.Close
Set rs = Nothing
Set oConect = Nothing
End Function

Public Function GetDetalleHabDevCajero(ByVal sAreaCod As String, ByVal sAgecod As String, _
        ByVal dFecha As Date, Optional ByVal sCajero As String = "", Optional nMoneda As Moneda = gMonedaNacional) As ADODB.Recordset


Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sFiltro As String

Set oCon = New COMConecta.DCOMConecta
Set rs = New Recordset
If oCon.AbreConexion = False Then Exit Function
sFiltro = ""
If sCajero <> "" Then
    sFiltro = " AND (H.cUsuDest = '" & sCajero & "' Or H.cUsuOrig = '" & sCajero & "') "
End If

sql = "SELECT SUBSTRING(M.cMovNro,9,2) + ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) Fecha, " _
    & "O.cOpeDesc, H.cUsuOrig, H.cUsuDest, ISNULL(RH.cPersNombre,'BOVEDA') cPersNombre, cSimbolo = CASE nMoneda WHEN 1 THEN 'S/.' ELSE 'US$' END, " _
    & "H.nMovImporte, M.nMovNro, H.nMoneda, M.cMovDesc FROM OpeTpo O INNER JOIN Mov M INNER JOIN MovHabDevCajero H " _
    & "LEFT JOIN (Select RH.cUser, P.cPersCod, P.cPersNombre From RRHH RH INNER JOIN Persona P ON RH.cPersCod = P.cPersCod) RH ON H.cUsuOrig = RH.cUser ON M.nMovNro = H.nMovNro " _
    & "ON O.cOpeCod = M.cOpeCod JOIN (Select MR.nMovNroRef From MovRef MR INNER JOIN Mov M1 ON MR.nMovNro = M1.nMovNro " _
    & "Where M1.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagModificado & ")) MR " _
    & "ON M.nMovNro = MR.nMovNroRef WHERE " _
    & "M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagModificado & ") " _
    & "AND H.cAreaCod = '" & sAreaCod & "' AND H.cAgeCod = '" & sAgecod & "' And H.nMoneda = " & nMoneda & " " _
    & "AND SUBSTRING(M.cMovNro,1,8) LIKE '" & Format$(dFecha, "yyyymmdd") & "' " & sFiltro & " " _
    & "ORDER BY M.cMovNro"
 
Set rs = oCon.CargaRecordSet(sql)
Set GetDetalleHabDevCajero = rs
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function YaRegistroEfectivo(ByVal sAgencia As String, ByVal dFecha As Date, _
            ByVal sUsuario As String, ByVal sOperacion As String) As Boolean
Dim sSQL As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

sSQL = "Select M.nMovNro From MovUserEfectivo E JOIN Mov M ON E.nMovNro = M.nMovNro Where " _
    & "M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And E.cUser = '" & sUsuario & "' " _
    & "And SUBSTRING(M.cMovNro,18,2) = '" & sAgencia & "' And M.cOpeCod IN ('" & sOperacion & "')"

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
Set rs = oConect.CargaRecordSet(sSQL)
If rs.EOF And rs.BOF Then
    YaRegistroEfectivo = False
Else
    YaRegistroEfectivo = True
End If
rs.Close
Set rs = Nothing
oConect.CierraConexion
Set oConect = Nothing
End Function

Public Function CierreOperaciones(ByVal nTipo As Integer, Optional sUsuario As String = "") As Boolean

End Function

'FONCODES
'CRSF - 24/06
Public Function GetMovFoncodes(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                              ByVal psCodUser As String, ByVal psCodAge As String) As ADODB.Recordset
Dim sSQL As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
If psCodUser <> "" Then
    lsFiltroUser = "  AND RIGHT(M.cMovNro,4)='" & psCodUser & "' "
End If

sSQL = "SELECT CONVERT(Varchar(10),CONVERT(DATETIME, SUBSTRING(M.cMovNro,1,8)),103) + ' ' + " _
    & "SUBSTRING(M.cMovNro,9,2) + ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) Fecha, " _
    & "O.cOpeDesc, RIGHT(M.cMovNro,4) Usu, F.cctacod, cSimbolo = 'S/.', MS.nMonto, M.nMovNro, " _
    & "nMoneda = 1, M.cMovDesc FROM OpeTpo O JOIN Mov M JOIN MovServicios MS ON MS.nMovNro = M.nMovNro " _
    & "ON O.cOpeCod = M.cOpeCod INNER JOIN FONCODESMOV FM ON FM.nNroMov = MS.nMovNro " _
    & " INNER JOIN FONCODES F ON FM.cCtaCod = F.cCtaCod WHERE M.nMovFlag = " & gMovFlagVigente & " AND " _
    & "M.cOpeCod = '" & psOpeCod & "' And SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " & lsFiltroUser _
    & "AND SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' And '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "ORDER BY M.nMovNro "
 
Set rs = oConect.CargaRecordSet(sSQL)
Set GetMovFoncodes = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
'PLAN BICI
'CRSF - 24/06
Public Function GetMovPlanBici(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                              ByVal psCodUser As String, ByVal psCodAge As String) As ADODB.Recordset
Dim sSQL As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
If psCodUser <> "" Then
    lsFiltroUser = "  AND RIGHT(M.cMovNro,4)='" & psCodUser & "' "
End If

sSQL = "SELECT CONVERT(Varchar(10),CONVERT(DATETIME, SUBSTRING(M.cMovNro,1,8)),103) + ' ' + " _
    & "SUBSTRING(M.cMovNro,9,2) + ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) Fecha, " _
    & "O.cOpeDesc, RIGHT(M.cMovNro,4) Usu, F.cctacod, cSimbolo = 'S/.', MS.nMonto, M.nMovNro, " _
    & "nMoneda = 1, M.cMovDesc FROM OpeTpo O JOIN Mov M JOIN MovServicios MS ON MS.nMovNro = M.nMovNro " _
    & "ON O.cOpeCod = M.cOpeCod INNER JOIN PLANBICIMOV FM ON FM.nNroMov = MS.nMovNro " _
    & " INNER JOIN PLANBICI F ON FM.cCtaCod = F.cCtaCod WHERE M.nMovFlag = " & gMovFlagVigente & " AND " _
    & "M.cOpeCod = '" & psOpeCod & "' And SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " & lsFiltroUser _
    & "AND SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' And '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "ORDER BY M.nMovNro "
 
Set rs = oConect.CargaRecordSet(sSQL)
Set GetMovPlanBici = rs
oConect.CierraConexion: Set oConect = Nothing
End Function


Public Function GetOpeIngEgreOpeCMACs(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, ByVal psCodAge As String, _
            ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String) As ADODB.Recordset

Dim sSQL As String
Dim rs As ADODB.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim oConect As COMConecta.DCOMConecta
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

Dim lsTablaDiaria As String

If pdFecha = pdFechaDia Then
    lsTablaDiaria = " MovDiario "
Else
    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
End If
    
    
Set rs = New ADODB.Recordset
If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.cMovNro,4)  = '" & psCodUser & "' "
End If



sSQL = "SELECT (OpeCol.cCodAge + '2') As nAgeOrd, OpeCol.cCodAge, cGrupoNombre, SUM(nTotalMov) AS nTotalMov, " _
    & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
    & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END ) AS Cheque, " _
    & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, G.cGrupoCod " _
    & "FROM GrupoOpe G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
    & "JOIN (SELECT MC.cPersCod As cCodAge, M.cOpeCod, O.cOpeDesc, COUNT(*) AS nTotalMov, " _
    & "  CASE WHEN MD.nDocTpo IS NULL THEN SUM(MC.nMonto) ELSE 0 END AS Efectivo, " _
    & "  CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN  SUM(MC.nMonto) ELSE 0 END AS Cheque, " _
    & "  CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN SUM(MC.nMonto) ELSE 0 END AS OrdenPago " _
    & "  FROM " & lsTablaDiaria & " M JOIN MovCMAC MC ON MC.nMovNro = M.nMovNro " _
    & "  LEFT JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
    & "  JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
    & "  WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
    & "  AND MC.nMoneda = '" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
    & "  GROUP BY M.cOpeCod, O.cOpeDesc, MD.nDocTpo, MC.cPersCod) As OpeCol  " _
    & "ON OpeCol.cOpeCod = GO.cOpeCod " _
    & "WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%' And G.nEfectivo = 1 " _
    & "GROUP BY OpeCol.cCodAge, cGrupoNombre, G.cGrupoCod "
sSQL = sSQL + " UNION " _
    & "(SELECT DISTINCT (P.cPersCod + '1') As nAgeOrd, P.cPersCod As cAgeCod, P.cPersNombre As cAgeDescripcion,0,0,0,0,'' " _
    & "FROM Persona P JOIN " _
    & "(SELECT MC.cPersCod As cCodAge " _
    & " FROM MovCMAC MC JOIN " & lsTablaDiaria & " M ON M.nMovNro = MC.nMovNro " _
    & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
    & " AND MC.nMoneda = '" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
    & ") Mov ON Mov.cCodAge = P.cPersCod) " _
    & " ORDER BY nAgeOrd "

Set rs = oConect.CargaRecordSet(sSQL)
Set GetOpeIngEgreOpeCMACs = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GrabaCierreAgencia(ByVal sMovNro As String, ByVal sOpeCod As String, ByVal sMovDesc As String) As Integer
Dim oMOv As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMOv = New COMDMov.DCOMMov

On Error GoTo Error
oMOv.BeginTrans
lbTrans = True
GrabaCierreAgencia = 1
oMOv.InsertaMov sMovNro, sOpeCod, sMovDesc, gMovEstContabNoContable, gMovFlagVigente
oMOv.CommitTrans
lbTrans = False
GrabaCierreAgencia = 0
Set oMOv = Nothing
Exit Function
Error:
    If lbTrans Then
        oMOv.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function YaRealizoCierreAgencia(ByVal sCodage As String, ByVal dFecha As Date) As Boolean
Dim sSQL As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

sSQL = "Select M.cOpeCod From Mov M " _
    & "Where M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And M.cOpeCod = '" & gOpeCajaCierreAgencia & "'" _
    & "And M.nMovFlag IN (" & gMovFlagVigente & ") AND Substring(M.cMovNro, 18,2) = '" & sCodage & "'"

Set rs = oConect.CargaRecordSet(sSQL)
If rs.EOF And rs.BOF Then
    YaRealizoCierreAgencia = False
Else
    YaRealizoCierreAgencia = True
End If
rs.Close
Set rs = Nothing
Set oConect = Nothing
End Function

Public Function YaRegistroEfectivoAutomatico(ByVal sCodage As String, ByVal dFecha As Date) As Integer
Dim sSQL As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

sSQL = "select cCodAge From EfectivoAutomatico "
sSQL = sSQL & " where convert(varchar(8), dfecha, 112)='" & Format(dFecha, "YYYYMMdd") & "' "
sSQL = sSQL & " And cCodAge='" & Trim(sCodage) & "' and nFlag IN (" & gMovFlagVigente & ")"

Set rs = oConect.CargaRecordSet(sSQL)
If rs.EOF And rs.BOF Then
    
    rs.Close
    
    sSQL = "Select M.cOpeCod From Mov M " _
    & "Where M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And M.cOpeCod = '" & gOpeCajaCierreAgencia & "'" _
    & "And M.nMovFlag IN (" & gMovFlagVigente & ") AND Substring(M.cMovNro, 18,2) = '" & sCodage & "'"

    Set rs = oConect.CargaRecordSet(sSQL)
    If rs.EOF And rs.BOF Then
        
        rs.Close
    
        sSQL = "Select M.* from mov M Inner Join MovUserEfectivo MU On M.nMovNro=MU.nMovNro "
        sSQL = sSQL & " Where M.nMovFlag IN (" & gMovFlagVigente & ") AND Substring(M.cMovNro, 18,2) = '" & sCodage & "' "
        sSQL = sSQL & " And M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And M.cOpeCod IN('" & gOpeBoveAgeRegEfect & "', '" & gOpeHabCajRegEfect & "') "
        
        Set rs = oConect.CargaRecordSet(sSQL)
        If rs.EOF And rs.BOF Then
            YaRegistroEfectivoAutomatico = 0 'Apto para registrar
        Else
            YaRegistroEfectivoAutomatico = 3 'Alguien ha registrado su efectivo
        End If
    Else
        YaRegistroEfectivoAutomatico = 2 'Ya efectuo cierre de agencia
    End If
Else
    YaRegistroEfectivoAutomatico = 1 'Ya registro manual
End If
rs.Close
Set rs = Nothing
Set oConect = Nothing

End Function

Public Sub RegistraEfectivoAutomatico(ByVal sCodage As String, ByVal dFecha As Date, ByVal sCodUsu As String, ByVal psNomAge As String)
Dim oConect As COMConecta.DCOMConecta
Dim oMOv As New COMDMov.DCOMMov
Dim clsMov As COMNContabilidad.NCOMContFunciones
Dim reg As ADODB.Recordset
Dim sSQL As String

Dim dfechaAnt As Date
Dim lnTempo As Long
Dim sMovNro As String
Dim lnMovNro As Long
Dim lbTrans As Boolean
Dim lsCodUser As String

On Error GoTo ErrorRegistraEfectivoAutomatico

lbTrans = True

dfechaAnt = DateAdd("d", -1, dFecha)

Set oConect = New COMConecta.DCOMConecta
Set clsMov = New COMNContabilidad.NCOMContFunciones

Set oMOv = New COMDMov.DCOMMov

If oConect.AbreConexion = False Then Exit Sub

'sSql = "select distinct mu.nmovnro, m.cOpeCod,mu.cUser, mu.cEfectivoCod, mu.nMonto "
'sSql = sSql & " from Mov M "
'sSql = sSql & " Inner Join MovUserEfectivo MU "
'sSql = sSql & " on M.nMovNro=MU.nMovNro "
'sSql = sSql & " Where cOpeCod in('" & gOpeBoveAgeRegEfect & "', '" & gOpeHabCajRegEfect & "') "
'sSql = sSql & " and m.cmovnro like '" & Format(dfechaAnt, "YYYYMMdd") & "%' "
'sSql = sSql & " and m.nmovflag IN(" & gMovFlagVigente & ") AND Substring(M.cMovNro, 18,2) = '" & sCodAge & "' "
'sSql = sSql & " Order by mu.nmovnro, m.copecod "

sSQL = "select distinct mu.cEfectivoCod, sum(mu.nMonto) as nMonto "
sSQL = sSQL & " from Mov M "
sSQL = sSQL & " Inner Join MovUserEfectivo MU "
sSQL = sSQL & " on M.nMovNro=MU.nMovNro "
sSQL = sSQL & " Where cOpeCod in('" & gOpeBoveAgeRegEfect & "', '" & gOpeHabCajRegEfect & "') "
sSQL = sSQL & " and m.cmovnro like '" & Format(dfechaAnt, "YYYYMMdd") & "%' "
sSQL = sSQL & " and m.nmovflag IN(" & gMovFlagVigente & ") AND Substring(M.cMovNro, 18,2) = '" & sCodage & "' "
sSQL = sSQL & " GROUP BY mu.cEfectivoCod "

Set reg = oConect.CargaRecordSet(sSQL)
oConect.CierraConexion
Set oConect = Nothing
If Not reg.EOF And Not reg.BOF Then
    lnTempo = 0
    oMOv.BeginTrans
    
    sMovNro = clsMov.GeneraMovNro(dFecha, sCodage, "SIST")
    
    oMOv.InsertaMov sMovNro, gOpeBoveAgeRegEfect, "REGISTRO DE EFECTIVO AUTOMATICO - " & sCodUsu, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oMOv.GetnMovNro(sMovNro)
    lsCodUser = "BOVE"
    
    Do While Not reg.EOF
        oMOv.InsertaMovUserEfectivo lnMovNro, lsCodUser, reg!cEfectivoCod, reg!nMonto
        Sleep 120
        'If reg.Bookmark = 18 Then Stop
        reg.MoveNext
    Loop
    oMOv.InsertaEfectivoAutomatico dFecha, sCodage, sCodUsu
    '****************** Cierre de agencia EJRS ******************************
    sMovNro = clsMov.GeneraMovNro(dFecha, sCodage, "SIST")
    oMOv.InsertaMov sMovNro, gOpeCajaCierreAgencia, "Cierre Caja Agencia Automatico : " & sCodage & " - " & psNomAge, gMovEstContabNoContable, gMovFlagVigente
    '************************************************************************
    oMOv.CommitTrans
End If
reg.Close
Set reg = Nothing

lbTrans = False

Set clsMov = Nothing
Set oMOv = Nothing

Exit Sub
ErrorRegistraEfectivoAutomatico:
    If lbTrans Then
        lbTrans = False
        oMOv.RollbackTrans
    End If
    Err.Raise vbObjectError + 100, "RegistraEfectivoAutomatico", "Efectivo Automatico "

End Sub

Public Sub RegistraExtornoAutomatico(ByVal sCodage As String, ByVal dFecha As Date, ByVal sCodUsu As String)
Dim oConect As COMConecta.DCOMConecta
Dim reg As ADODB.Recordset
Dim sSQL1 As String
Dim sSql2 As String

Dim lbTrans As Boolean

On Error GoTo ErrorExtornoEfectivoAutomatico
lbTrans = True

Set oConect = New COMConecta.DCOMConecta

If oConect.AbreConexion = False Then Exit Sub

sSQL1 = "Update Mov set nMovFlag=" & gMovFlagExtornado & " "
sSQL1 = sSQL1 & " Where cOpeCod in('" & gOpeBoveAgeRegEfect & "', '" & gOpeHabCajRegEfect & "') "
sSQL1 = sSQL1 & " and cmovnro like '" & Format(dFecha, "YYYYMMdd") & "%' "
sSQL1 = sSQL1 & " and nmovflag IN(" & gMovFlagVigente & ") AND Substring(cMovNro, 18,2) = '" & sCodage & "' "

sSql2 = " Update EfectivoAutomatico set nFlag=" & gMovFlagExtornado & " "
sSql2 = sSql2 & " where convert(varchar(8), dfecha, 112)='" & Format(dFecha, "YYYYMMdd") & "' and cCodAge='" & Trim(sCodage) & "' "
sSql2 = sSql2 & " and nflag IN(" & gMovFlagVigente & ")"

oConect.BeginTrans
oConect.Ejecutar sSQL1
oConect.Ejecutar sSql2
oConect.CommitTrans

oConect.CierraConexion
Set oConect = Nothing

lbTrans = False

Exit Sub
ErrorExtornoEfectivoAutomatico:
    If lbTrans Then
        lbTrans = False
        oConect.RollbackTrans
    End If
    Err.Raise vbObjectError + 100, "ExtornaEfectivoAutomatico", "Efectivo Automatico "

End Sub


Public Function GetUsuariosMovEfectivo(ByVal sAgencia As String, ByVal dFecha As Date) As ADODB.Recordset
Dim sSQL As String
Dim oConect As COMConecta.DCOMConecta

sSQL = "Select RIGHT(M.cMovNro,4) Usuario, P.cPersNombre From Persona P JOIN RRHH RH JOIN Mov M JOIN " _
    & "(SELECT GO.cOpeCod From GrupoOpe G JOIN GruposOpe GO ON G.cGrupoCod = GO.cGrupoCod " _
    & "WHERE G.nEfectivo = 1 And cOpeCod NOT LIKE '9010%' And cOpeCod NOT IN (401401, 402401) " _
    & "And cOpeCod NOT IN (Select cOpeCod from OpeDoc Where nDocTpo = 47) ) O ON " _
    & "M.cOpeCod = O.cOpeCod ON RH.cUser = RIGHT(M.cMovNro,4) ON P.cPersCod = RH.cPersCod " _
    & "Where M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' " _
    & "And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & sAgencia & "' " _
    & "GROUP BY RIGHT(M.cMovNro,4), P.cPersNombre"


Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
Set GetUsuariosMovEfectivo = oConect.CargaRecordSet(sSQL)
Set oConect = Nothing
End Function
Public Function GetCajeroOpeIngxRef(ByVal psPersCod As String, ByVal pnMoneda As Moneda) As ADODB.Recordset
Dim sql As String
Dim oCon As COMConecta.DCOMConecta

sql = "SELECT  SUBSTRING(M.CMOVNRO,7,2) + '/' + SUBSTRING(M.CMOVNRO,5,2) + '/' + SUBSTRING(M.CMOVNRO,1,4) + ' ' + SUBSTRING(M.CMOVNRO,9,2) + ':' +SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) AS FECHA  , " _
    & "         m.cMovDesc , MV.NMOVIMPORTE, m.nMovNro, MV.nMoneda " _
    & " FROM    MOV M " _
    & "         JOIN MOVOPEVARIAS MV ON MV.NMOVNRO = M.NMOVNRO " _
    & "         JOIN MOVGASTO MG ON MG.NMOVNRO = M.NMOVNRO " _
    & "         JOIN PERSONA P ON P.CPERSCOD = MG.CPERSCOD " _
    & " WHERE   M.copecod = 300405 AND MG.CPERSCOD = '" & psPersCod & "' " _
    & "         AND M.NMOVFLAG = " & gMovFlagVigente & " AND MV.NMONEDA =" & pnMoneda & "" _
    & "         AND NOT EXISTS ( SELECT * FROM MOV M1 JOIN MOVREF MR ON MR.NMOVNRO = M1.NMOVNRO WHERE MR.NMOVNROREF = M.NMOVNRO AND M1.NMOVFLAG = " & gMovFlagVigente & ")"

Set oCon = New COMConecta.DCOMConecta
If oCon.AbreConexion = False Then Exit Function
Set GetCajeroOpeIngxRef = oCon.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function GetSobFalMayUNO(ByVal psFecha As String, ByVal psCodAge As String) As ADODB.Recordset
Dim sql As String
Dim oCon As COMConecta.DCOMConecta

sql = "SELECT  RIGHT(M.CMOVNRO,4) AS CUSER, m.cmovnro, M.cOpeCod, nMoneda ,  nMovImporte, O.cOpeDesc, cSimbolo = CASE nMoneda WHEN 1 THEN 'S/.' ELSE 'US$' END " _
    & " FROM    MOV M " _
    & "         JOIN MOVOPEVARIAS MV ON MV.NMOVNRO =M.NMOVNRO " _
    & "         JOIN OPETPO O ON O.COPECOD = M.COPECOD " _
    & " WHERE   LEFT(M.CMOVNRO,8)='" & Format(psFecha, "yyyymmdd") & "' AND  M.COPECOD IN ('901009','901010','901019','901020') " _
    & "         AND SubString(M.cMovNro,18,2)='" & psCodAge & "' and abs(nMovImporte)>1 AND M.NMOVFLAG = " & gMovFlagVigente & ""

Set oCon = New COMConecta.DCOMConecta
If oCon.AbreConexion = False Then Exit Function
Set GetSobFalMayUNO = oCon.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing
End Function


