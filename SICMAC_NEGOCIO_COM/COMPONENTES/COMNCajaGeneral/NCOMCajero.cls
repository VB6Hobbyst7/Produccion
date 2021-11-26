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
Public oImp As COMFunciones.FCOMVarImpresion
Public gcPEN As New COMDConstantes.DCOMConstantes 'marg ers044-2016
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub IniciaImpresora(Optional ByVal nImpresora As COMFunciones.Impresoras = gEPSON)
Set oImp = New COMFunciones.FCOMVarImpresion
oImp.Inicia nImpresora
End Sub

Private Sub Class_Initialize()
Dim oIni As COMConecta.DCOMClasIni
Set oIni = New COMConecta.DCOMClasIni

IniciaImpresora
vsBaseComunes = oIni.BaseComunes
vsBasePesonas = oIni.BasePersonas
Dim oCons As COMDConstSistema.NCOMConstSistema
Set oCons = New COMDConstSistema.NCOMConstSistema

'Ya no se debe asignar en esta DLL
'gsCodCMAC = oCons.LeeConstSistema(COMDConstantes.gConstSistCodCMAC)

Set oIni = Nothing
Set oCons = Nothing
End Sub
Public Function GetCierreAgencias(ByVal pdFecha As Date, ByVal sOpeCod As String, ByVal sMovDesc As String, ByVal psCodAge As String) As adodb.Recordset
Dim oConect As COMConecta.DCOMConecta
Dim lsTablaDiaria As String
Dim sql As String
Dim rs As New adodb.Recordset

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

Dim oMov As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New COMDMov.DCOMMov

On Error GoTo GrabaHabilitaAgenciaErr
oMov.BeginTrans
lbTrans = True
GrabaHabilitaAgencia = 1
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovHabDev lnMovNro, pnImporte, psAreaCod, psAgeCod, nMoneda, sUsuOrig, sUsuDest
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

'MADM 20110125
Public Function GrabaHabilitaAgenciaPreCuadre(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
            ByVal pnImporte As Double, ByVal psAreaCod As String, ByVal psAgeCod As String, _
            ByVal nMoneda As Moneda, ByVal sUsuOrig As String, ByVal sUsuDest As String) As Integer

Dim oMov As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New COMDMov.DCOMMov

On Error GoTo GrabaHabilitaAgenciaPreCuadreErr
oMov.BeginTrans
lbTrans = True
GrabaHabilitaAgenciaPreCuadre = 1
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovHabDevPreCuadre lnMovNro, pnImporte, psAreaCod, psAgeCod, nMoneda, sUsuOrig, sUsuDest
oMov.CommitTrans
lbTrans = False
GrabaHabilitaAgenciaPreCuadre = 0
Set oMov = Nothing
Exit Function
GrabaHabilitaAgenciaPreCuadreErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'END MADM

Public Function GetMovHabBovAgeCajero(ByVal psOpeCod As String, ByVal psAreaCod As String, _
                ByVal psAgeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                Optional ByVal psCodUser As String = "") As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltro As String, sFiltroArea As String

Set oCon = New COMConecta.DCOMConecta
Set rs = New adodb.Recordset
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
                Optional ByVal psCodUser As String = "") As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltro As String

Set oCon = New COMConecta.DCOMConecta
Set rs = New adodb.Recordset
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
Dim sSql As String, rs As adodb.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New adodb.Recordset
If oCon.AbreConexion = False Then Exit Function

GetBuscarMov = False

    sSql = "SELECT *    FROM MOVHABDEVCAJERO MH "
    sSql = sSql & "     JOIN MOV M ON M.NMOVNRO=MH.NMOVNRO "
    sSql = sSql & "     WHERE CusuORIG='" & psCodUser & "' and MH.cAgeCod='" & psCodAge & "' "
    sSql = sSql & "  AND M.CMOVNRO LIKE '" & psFecha & "%" & psCodUser & "' and M.copecod='" & psOpeCod & "' and m.nmovflag=0 and mh.nmoneda=" & pnMoneda
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF Then
        GetBuscarMov = True
    End If
    
    Set rs = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
    

End Function


Public Function GetMovHabCajDevABove(ByVal psOpeCod As String, ByVal psAreaCod As String, _
                ByVal psAgeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date) As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltro As String

Set oCon = New COMConecta.DCOMConecta
Set rs = New adodb.Recordset
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
Dim rs As adodb.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New adodb.Recordset

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
    GetMovUserBilletaje = IIf(pbnMovNro, rs!nMovNro, rs!cMovNro)
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion
Set oCon = Nothing

End Function

Public Function GrabaRegistroEfectivo(ByVal psFormatoFecha As String, _
            ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
            ByVal rsBillIngMN As adodb.Recordset, ByVal rsBillIngME As adodb.Recordset, _
            ByVal rsMonIngMN As adodb.Recordset, ByVal rsMonIngME As adodb.Recordset, _
            ByVal psCodUser As String, ByVal pbNuevoMov As Boolean, Optional ByRef pnMovNroRegEfec As Long) As Integer
Dim oMov As COMDMov.DCOMMov
Set oMov = New COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
On Error GoTo GrabaRegistroEfectivoErr

GrabaRegistroEfectivo = 1
oMov.BeginTrans
lbTrans = True
If pbNuevoMov Then
    oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(psMovNro)
Else
    lnMovNro = oMov.GetnMovNro(psMovNro)
End If
pnMovNroRegEfec = lnMovNro '**DAOR 20080125
'Moneda Nacional
If Not rsBillIngMN Is Nothing Then 'Billetes
    If rsBillIngMN.State = adStateOpen Then
        If Not rsBillIngMN.EOF And Not rsBillIngMN.BOF Then
            Do While Not rsBillIngMN.EOF
                If pbNuevoMov Then
                    oMov.InsertaMovUserEfectivo lnMovNro, psCodUser, rsBillIngMN("cEfectivoCod"), rsBillIngMN("Monto")
                Else
                    oMov.ActualizaMovUserEfectivo lnMovNro, psCodUser, rsBillIngMN("cEfectivoCod"), rsBillIngMN("Monto")
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
                    oMov.InsertaMovUserEfectivo lnMovNro, psCodUser, rsMonIngMN("cEfectivoCod"), rsMonIngMN("Monto")
                Else
                    oMov.ActualizaMovUserEfectivo lnMovNro, psCodUser, rsMonIngMN("cEfectivoCod"), rsMonIngMN("Monto")
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
                    oMov.InsertaMovUserEfectivo lnMovNro, psCodUser, rsBillIngME("cEfectivoCod"), rsBillIngME("Monto")
                Else
                    oMov.ActualizaMovUserEfectivo lnMovNro, psCodUser, rsBillIngME("cEfectivoCod"), rsBillIngME("Monto")
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
                    oMov.InsertaMovUserEfectivo lnMovNro, psCodUser, rsMonIngME("cEfectivoCod"), rsMonIngME("Monto")
                Else
                    oMov.ActualizaMovUserEfectivo lnMovNro, psCodUser, rsMonIngME("cEfectivoCod"), rsMonIngME("Monto")
                End If
                rsMonIngME.MoveNext
            Loop
        End If
        rsMonIngME.Close: Set rsMonIngME = Nothing
    End If
End If
oMov.CommitTrans
lbTrans = False

GrabaRegistroEfectivo = 0
Set oMov = Nothing

Exit Function
GrabaRegistroEfectivoErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaRegistroEfectivo", Err.Description & " Error Registro de Efectivo Cajero"
End Function

'MADM 20110115
Public Function GrabaRegistroEfectivoCorte(ByVal psFormatoFecha As String, _
            ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
            ByVal pnTotSol As Double, ByVal pnTotDol As Double, _
            ByVal psCodUser As String, ByVal pbNuevoMov As Boolean, Optional ByRef pnMovNroRegEfec As Long) As Integer
Dim oMov As COMDMov.DCOMMov
Set oMov = New COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
On Error GoTo GrabaRegistroEfectivoCorteErr

GrabaRegistroEfectivoCorte = 1
oMov.BeginTrans
lbTrans = True
If pbNuevoMov Then
    oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(psMovNro)
Else
    lnMovNro = oMov.GetnMovNro(psMovNro)
End If
pnMovNroRegEfec = lnMovNro '**DAOR 20080125

oMov.InsertaMovPreCuadre lnMovNro, psCodUser, pnTotSol, pnTotDol
 
oMov.CommitTrans
lbTrans = False

GrabaRegistroEfectivoCorte = 0
Set oMov = Nothing

Exit Function
GrabaRegistroEfectivoCorteErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaRegistroEfectivo", Err.Description & " Error Registro de Efectivo Cajero"
End Function

Public Function GetValidacionRegistroPreCuadre(ByVal psCodUser As String, ByVal psCodAge As String, ByVal psFecha As String, ByVal psOpeCod As String) As adodb.Recordset
Dim loCon As COMConecta.DCOMConecta
Dim lsSql As String
Dim lRs As adodb.Recordset
On Error GoTo GetValidacionRegistroPreCuadreErr
    Set loCon = New COMConecta.DCOMConecta
        loCon.AbreConexion
    lsSql = " exec stp_sel_ValidacionRegistroPreCuadre '" & psCodUser & "','" & psCodAge & "','" & psFecha & "','" & psOpeCod & "'"
    Set lRs = loCon.CargaRecordSet(lsSql)
    Set GetValidacionRegistroPreCuadre = lRs
    loCon.CierraConexion
    Exit Function
GetValidacionRegistroPreCuadreErr:
    Set GetValidacionRegistroPreCuadre = Nothing
    loCon.CierraConexion
End Function

Public Function ObtenerValidacionRegistroPreCuadre(ByVal psCodUser As String, ByVal psCodAge As String, ByVal psFecha As String, ByVal psOpeCod As String) As Boolean
Dim loCon As COMConecta.DCOMConecta
Dim lsSql As String
Dim lRs As adodb.Recordset

    Set loCon = New COMConecta.DCOMConecta
        loCon.AbreConexion
    lsSql = " exec stp_sel_ValidacionRegistroPreCuadre '" & psCodUser & "','" & psCodAge & "','" & psFecha & "','" & psOpeCod & "'"
    Set lRs = loCon.CargaRecordSet(lsSql)
    
    If lRs.EOF And lRs.BOF Then
        ObtenerValidacionRegistroPreCuadre = False
    Else
        ObtenerValidacionRegistroPreCuadre = True
    End If
    
    lRs.Close
    Set lRs = Nothing
    loCon.CierraConexion
    Exit Function

    loCon.CierraConexion
End Function

'END MADM

Public Function GetBilletajeCajero(ByVal psMovNro As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, Optional ByVal psDenominacion As String = "") As adodb.Recordset
Dim sql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset
Dim lsFiltroDenom As String

Set rs = New adodb.Recordset
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
                                    ByVal rsBillIngreso As adodb.Recordset, ByVal rsMonIng As adodb.Recordset, _
                                    ByVal psAreaCod As String, ByVal psAgeCod As String, _
                                    ByVal psPersCodUser As String, ByVal pnTotal As Currency, _
                                    ByVal pnMovNroReg As Long) As Integer
            
Dim oMov As COMDMov.DCOMMov
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New COMDMov.DCOMMov
On Error GoTo GrabaDevolucionABovedaErr

lnMovItem = 0: lnMovOrden = 0

GrabaDevolucionABoveda = 1
oMov.BeginTrans
lbTrans = True
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)
'guardamos la cuenta en el debe
If Not rsBillIngreso Is Nothing Then
    If rsBillIngreso.State = adStateOpen Then
        If Not rsBillIngreso.EOF And Not rsBillIngreso.BOF Then
            Do While Not rsBillIngreso.EOF
                If rsBillIngreso!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, rsBillIngreso!Monto
                    lnMovOrden = lnMovOrden + 1
                    oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, rsBillIngreso!cEfectivoCod
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
                    oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, rsMonIng!Monto
                    lnMovOrden = lnMovOrden + 1
                    oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, rsMonIng!cEfectivoCod
                End If
                rsMonIng.MoveNext
            Loop
        End If
        rsMonIng.Close: Set rsMonIng = Nothing
    End If
End If
lnMovOrden = lnMovOrden + 1
oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
oMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, psAgeCod, psAreaCod


lnMovItem = lnMovItem + 1: lnMovOrden = 0
oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnTotal * -1
lnMovOrden = lnMovOrden + 1
oMov.InsertaMovGasto lnMovNro, psPersCodUser, ""

oMov.InsertaMovRef lnMovNro, pnMovNroReg

oMov.CommitTrans
lbTrans = False
GrabaDevolucionABoveda = 0

Set oMov = Nothing

Exit Function
GrabaDevolucionABovedaErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaDevolucionABoveda", Err.Description & " Error Devolucion de Efectivo de Cajero a Boveda"
End Function
Public Function GetMovUserBilletajeDevuelto(ByVal psOpeCod As String, ByVal psCodUser As String, ByVal pdFecSis As Date) As String
Dim sql As String
Dim rs As adodb.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New adodb.Recordset

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

Dim oMov As COMDMov.DCOMMov
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim lbTrans As Boolean
Dim lnMovNro As Long


Set oMov = New COMDMov.DCOMMov
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
 
'oMov.InsertaMovHabDev lnMovNro, psUserDest, pnImporte

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

Public Function GrabaConfHabilitaAgencia(ByVal sMovNro As String, ByVal sOpeCod As String, ByVal sMovDesc As String, _
            ByVal nMovNroHab As Long, Optional ByVal psMotExtorno As Variant) As Integer

Dim oMov As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim nMovNro As Long

Set oMov = New COMDMov.DCOMMov
oMov.BeginTrans
lbTrans = True
GrabaConfHabilitaAgencia = 1
oMov.InsertaMov sMovNro, sOpeCod, sMovDesc, gMovEstContabNoContable, gMovFlagVigente
nMovNro = oMov.GetnMovNro(sMovNro)

'***CTI3 (ferimoro) 17102018
'oMov.dInsertDetExtMov nMovNro, sOpeCod, psMotExtorno(0), psMotExtorno(1)

oMov.InsertaMovRef nMovNro, nMovNroHab
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
                                    ByVal psCodUser As String) As adodb.Recordset
Dim sql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset
Dim lsFiltroDenom As String
Dim lsFiltroUser As String
Set rs = New adodb.Recordset
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
                                    ByVal psCodUser As String) As adodb.Recordset
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset
Dim lsFiltroDenom As String
Dim lsFiltroUser As String
Dim sSql As String
Set rs = New adodb.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
lsFiltroUser = ""

'RIRO20170707 Convertido a Procedimiento Almacenado
'If psCodUser <> "" Then
'    lsFiltroUser = " AND MH.cUsuOrig = '" & psCodUser & "' "
'End If
'
'sSql = "Select CONVERT(Varchar(12),CONVERT(DATETIME,SUBSTRING(M.cMovNro,1,8)),103) + ' ' + " _
'    & "SUBSTRING(M.cMovNro,9,2)+ ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) AS Fecha, " _
'    & "O.cOpeDesc, MH.cUsuOrig, ISNULL(P.cPersNombre,'BOVEDA') cPersNombre, " _
'    & "Moneda = CASE WHEN MH.nMoneda = 1 THEN 'S/.' ELSE 'US$' END, MH.nMovImporte, M.nMovNro, MH.nMoneda, M.cMovDesc " _
'    & "FROM OpeTpo O JOIN Mov M JOIN MovHabDevCajero MH LEFT JOIN RRHH RH JOIN Persona P ON RH.cPersCod = P.cPersCod ON " _
'    & "MH.cUsuOrig = RH.cUser ON M.nMovNro = MH.nMovNro ON O.cOpeCod = M.cOpeCod " _
'    & "WHERE M.cOpeCod = '" & psOpeCod & "' " & lsFiltroUser & " And Substring(M.cMovNro,1,8) " _
'    & "BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' And '" & Format(pdHasta, "yyyymmdd") & "' " _
'    & "And M.nMovFlag IN (" & gMovFlagVigente & ") " _
'    & "Order by M.nMovNro"
'END RIRO Fin de comentario

sSql = "stp_sel_listardevCajero '" & psOpeCod & "', '" & Format(pdDesde, "yyyymmdd") & "', '" & Format(pdHasta, "yyyymmdd") & "', '" & psCodUser & "'"

Set rs = oConect.CargaRecordSet(sSql)
Set GetDevCajero = rs
oConect.CierraConexion: Set oConect = Nothing
End Function

'MADM 20110201
Public Function GetDevCajeroPreCuadre(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                    ByVal psCodUser As String) As adodb.Recordset
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset
Dim lsFiltroDenom As String
Dim lsFiltroUser As String
Dim sSql As String
Set rs = New adodb.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
lsFiltroUser = ""
If psCodUser <> "" Then
    lsFiltroUser = " AND MH.cUsuOrig = '" & psCodUser & "' "
End If

sSql = "Select CONVERT(Varchar(12),CONVERT(DATETIME,SUBSTRING(M.cMovNro,1,8)),103) + ' ' + " _
    & "SUBSTRING(M.cMovNro,9,2)+ ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) AS Fecha, " _
    & "O.cOpeDesc, MH.cUsuOrig, ISNULL(P.cPersNombre,'BOVEDA') cPersNombre, " _
    & "Moneda = CASE WHEN MH.nMoneda = 1 THEN 'S/.' ELSE 'US$' END, MH.nMovImporte, M.nMovNro, MH.nMoneda, M.cMovDesc " _
    & "FROM OpeTpo O JOIN Mov M JOIN MovHabDevCajeroPreCuadre MH LEFT JOIN RRHH RH JOIN Persona P ON RH.cPersCod = P.cPersCod ON " _
    & "MH.cUsuOrig = RH.cUser ON M.nMovNro = MH.nMovNro ON O.cOpeCod = M.cOpeCod " _
    & "WHERE M.cOpeCod = '" & psOpeCod & "' " & lsFiltroUser & " And Substring(M.cMovNro,1,8) " _
    & "BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' And '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "And M.nMovFlag IN (" & gMovFlagVigente & ") " _
    & "Order by M.nMovNro"

Set rs = oConect.CargaRecordSet(sSql)
Set GetDevCajeroPreCuadre = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
'END MADM

'MADM 20110926 - Extorno Confirmacion devolucion a BOVEDA
Public Function GetDevCajeroConfDevolucionBove(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                    ByVal psCodUser As String) As adodb.Recordset
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset
Dim sSql As String

Set rs = New adodb.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
psOpeCod = "901012"

sSql = "exec stp_sel_CajeroConfDevolcionBove '" & psOpeCod & "','" & Format(pdDesde, "yyyymmdd") & "','" & Format(pdHasta, "yyyymmdd") & "','" & psCodUser & "'"

Set rs = oConect.CargaRecordSet(sSql)
Set GetDevCajeroConfDevolucionBove = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
'END MADM

Public Function GetTransfEntreCajeros(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                    ByVal psCodUser As String) As adodb.Recordset
Dim sql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset
Dim lsFiltroDenom As String
Dim lsFiltroUser As String

Set rs = New adodb.Recordset
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
Public Function GrabaCompraVenta(ByVal gdFecSis As String, ByVal psFormatoFecha As String, _
            ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
            ByVal pnImporte As Double, ByVal pnTipoCambio As Double, _
            ByVal psCodPers As String, Optional bLavDinero As Boolean = False, Optional SPersBen As String = "", _
            Optional psTitPersLavDinero As String = "", Optional psOrdPersLavDinero As String = "", _
            Optional psReaPersLavDinero As String = "", Optional psBenPersLavDinero As String = "", _
            Optional psVisPersLavDinero As String = "", _
            Optional pnMovNro As Long = 0, Optional ByVal pnTipoPago As Integer = 1, Optional ByVal psCtaCodCargo As String = "", _
            Optional ByRef pMatAho As Variant = "", _
            Optional ByVal pbITFAplica As Boolean, _
            Optional ByVal pnMontoITF As Currency, _
            Optional ByVal pnMontoCargoCta As Currency = 0#, _
            Optional pnMovNroTransfer As Long = -1, Optional pnMonedaTrans As Moneda = gMonedaNacional, _
            Optional ByVal pnMovNroRVD As Long = 0, _
            Optional ByVal pnMontoRVD As Currency = 0) As Integer
            'By Capi 28012008 se agrego 5 ult parametros
' JUCS 20170822, Segun TI-ERS 002-2017*****************************************************************************************************************
Dim RegTran As adodb.Recordset, RegOpeCta As adodb.Recordset, RegTmp As adodb.Recordset
Dim nMovSoles As Currency, nImporteSoles As Currency, TCF As Currency, CVDebe As Currency, CVHaber As Currency
Dim nCtaDeb1 As Currency, nCtaHab1 As Currency, vCtaDolar As Currency, paAsiento As Currency 'JUCS 20171026
Dim paAsiento2 As Double
Dim psFecha As Date
Dim vAgencia As String, vCodConta As String, vNoCtaCnt As String, vParche As String, vAG As String, sSql As String ', tmpSql As String
'Dim TCC As Currency, TCV As Currency ', TCCPonderado As Currency
Dim lsCVMEGanacia As String, lsCVMEPerdida As String
Dim lsCajaSoles As String, lsCajaDolares As String, lsCajaAgenciaSoles As String, lsCajaAgenciaDolares As String, FechaHora As String
Dim lnItem As Long ', lnItemMe As Long
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
Dim oAsi As COMDCajaGeneral.DCOMAsiento
Set oAsi = New COMDCajaGeneral.DCOMAsiento
'*****************************************************************************************************************************************************************
            Dim oMov As COMDMov.DCOMMov
            Dim lbTrans As Boolean
            Dim lnMovNro As Long
            Set oMov = New COMDMov.DCOMMov
            Dim objTC As COMDConstSistema.DCOMTipoCambio
            Set objTC = New COMDConstSistema.DCOMTipoCambio
            Dim oRsTCambio As adodb.Recordset
            Set oRsTCambio = New adodb.Recordset

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

On Error GoTo GrabaCompraVentaErr
Dim lsOpeCodCargoCta As String 'CTI6 ERS0112020
Dim lsOpeCodCargoCtaDescripcion As String 'CTI6 ERS0112020
Dim lsOpeCod As String 'CTI6 ERS0112020

'Codigo de Operacion'
'CTI6 ERS112020
Select Case pnTipoPago
    Case gCVTipoPagoEfectivo   ' Efectivo
            lsOpeCod = psOpeCod
      Case gCVTipoPagoCargoCta
      If Mid(psCtaCodCargo, 9, 1) = gMonedaNacional Then
                lsOpeCodCargoCta = gAhoCargoCompra
                lsOpeCodCargoCtaDescripcion = "Cargo a Cuenta COMPRA: " & lsOpeCodCargoCta
                lsOpeCod = gCompraMECargoCtaAhorro
            Else
                lsOpeCodCargoCta = gAhoCargoVenta
                lsOpeCodCargoCtaDescripcion = "Cargo a Cuenta VENTA: " & lsOpeCodCargoCta
                lsOpeCod = gVentaMECargoCtaAhorro
            End If
       Case gCVTipoPagoVoucher
            If pnMonedaTrans = gMonedaNacional Then
                lsOpeCodCargoCtaDescripcion = "Voucher VENTA"
                lsOpeCod = gVentaMEVoucher
            Else
                lsOpeCodCargoCtaDescripcion = "Voucher COMPRA"
                lsOpeCod = gCompraMEVoucher
            End If
   End Select
'End CTI6
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
oMov.BeginTrans
lbTrans = True
GrabaCompraVenta = 1
'oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente 'Comentado CTI6 ERS0112020
oMov.InsertaMov psMovNro, lsOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente 'CTI6 ERS112020
lnMovNro = oMov.GetnMovNro(psMovNro)
'ALPA 20081010***************************************************************************************************
pnMovNro = lnMovNro
oMov.InsertaMovCompraVenta lnMovNro, psCodPers, pnImporte
oMov.InsertaMovTpoCambio lnMovNro, pnTipoCambio, lnValVent, lnValComp, lnValFijo, lnValFijoDia, lnValVentEsp, lnValCompEsp, lnValPond, lnValPondVenta, lnValPondREU, lnValSBSDia, lnValCompTr, lnValVentTr

If SPersBen = "" Then SPersBen = psCodPers

''If bLavDinero Then
'oMov.InsertaMovLavDinero lnMovNro, psTitPersLavDinero, psOrdPersLavDinero, psReaPersLavDinero, psBenPersLavDinero, psVisPersLavDinero, nTipoREU, nMontoAcu, sOrigen
'End If

' JUCS 20170822, Segun TI-ERS 002-2017*****************************************************************************************************************
FechaHora = CStr(Format(gdFecSis, "yyyy-MM-dd")) & " " & CStr(Format(Time, "hh:mm:ss.000"))
  '  gsCtaCodFoncodes = oAsi.GetAsientoParametro(1)
    lsCVMEGanacia = oAsi.GetAsientoParametro(2)
    lsCVMEPerdida = oAsi.GetAsientoParametro(3)
    lsCajaSoles = oAsi.GetAsientoParametro(4)
    lsCajaDolares = oAsi.GetAsientoParametro(5)
    lsCajaAgenciaSoles = oAsi.GetAsientoParametro(6)
    lsCajaAgenciaDolares = oAsi.GetAsientoParametro(7)

    TCF = lnValFijoDia
    nImporteSoles = pnImporte
    lnItem = 0
 
   'CTI6 ERS112020
    sSql = " SELECT cCtaContCod , cOpeCtaDH  FROM OpeCtaNeg  " _
             & " WHERE cOpeCod = '" & lsOpeCod & "'"

'   sSql = " SELECT cCtaContCod , cOpeCtaDH  FROM OpeCtaNeg  " _
'             & " WHERE cOpeCod = '" & psOpeCod & "'"
                       oCon.AbreConexion
                        Set RegOpeCta = oCon.CargaRecordSet(sSql)

                   If Not (RegOpeCta.BOF And RegOpeCta.EOF) Then
                        Do While Not RegOpeCta.EOF
                             vCodConta = Trim(RegOpeCta!cCtaContCod)
                             vAG = VarAG(Mid(psMovNro, 15, 5), vCodConta)
                             vCodConta = Replace(vCodConta, "AG", vAG, 1, 1, vbTextCompare)
                             vParche = AsientoParche(vCodConta, True, oCon.ConexionActiva)
                             vAgencia = Mid(psMovNro, 18, 2)
                             'Mejoras TI-ERS 002-2017 JUCS por problemas de redondeo
                             nMovSoles = Format(pnImporte * pnTipoCambio, "#,#0.00") 'Modificación JUCS 08062018 Acta 044-2018
                             'nMovSoles = pnImporte * pnTipoCambio
                             'nMovSoles = Format(nMovSoles, "#,#0.00")
                             paAsiento = pnImporte * TCF
                             paAsiento = Format(paAsiento, "#,#0.00")
                             'paAsiento = nImporteSoles * TCF
                            lnItem = lnItem + 1
                            'lnItemMe = lnItemMe + 1
                            'Cambia por la cuenta correcta
                             If Len(vParche) > 0 Then
                                 vCodConta = vParche
                             End If
                             If RegOpeCta!cOpeCtaDH = "D" Then
                                 If Left(vCodConta, 4) = lsCajaSoles Then
                                    oMov.InsertaMovCta lnMovNro, lnItem, vCodConta, nMovSoles
                                    'Call IngresarAsientoDN(FechaHora, vCodConta, nMovSoles, 0, 1, vAgencia, pnMovNro, psOpeCod, , oCon.ConexionActiva, TCF)
                                    Call IngresarAsientoDN(FechaHora, vCodConta, nMovSoles, 0, 1, vAgencia, pnMovNro, lsOpeCod, , oCon.ConexionActiva, TCF) 'CTI6 ERS0112020
                                    CVDebe = CVDebe + nMovSoles

                                ElseIf Left(vCodConta, 4) = lsCajaDolares Then 'Or Left(vCodConta, 4) = "2825"
                                    'CTI6 ERS112020
                                    Call IngresarAsientoDN(FechaHora, vCodConta, paAsiento, 0, 1, vAgencia, pnMovNro, lsOpeCod, , oCon.ConexionActiva, TCF)
                                    'Call IngresarAsientoDN(FechaHora, vCodConta, paAsiento, 0, 1, vAgencia, pnMovNro, psOpeCod, , oCon.ConexionActiva, TCF)'CTI6 ERS0112020
                                    oMov.InsertaMovCta lnMovNro, lnItem, vCodConta, paAsiento
                                    oMov.InsertaMovMe lnMovNro, lnItem, nImporteSoles
                                    lnItem = lnItem + 1
                                    'CTI6 ERS112020
                                    'Call IngresarAsientoDN(FechaHora, vCodConta, nImporteSoles, 0, 2, vAgencia, pnMovNro, psOpeCod, , oCon.ConexionActiva, TCF)
                                    Call IngresarAsientoDN(FechaHora, vCodConta, nImporteSoles, 0, 2, vAgencia, pnMovNro, lsOpeCod, , oCon.ConexionActiva, TCF)
                                    CVDebe = CVDebe + paAsiento
                                    
                               End If
                               
                            ElseIf RegOpeCta!cOpeCtaDH = "H" Then
                            nMovSoles = Abs(nMovSoles)
                            nImporteSoles = Abs(nImporteSoles)
                                If Left(vCodConta, 4) = lsCajaSoles Then
                                     'CTI6 ERS112020
                                     oMov.InsertaMovCta lnMovNro, lnItem, vCodConta, nMovSoles * -1
                                     'Call IngresarAsientoDN(FechaHora, vCodConta, 0, nMovSoles, 1, vAgencia, pnMovNro, psOpeCod, , oCon.ConexionActiva, TCF)
                                     Call IngresarAsientoDN(FechaHora, vCodConta, 0, nMovSoles, 1, vAgencia, pnMovNro, lsOpeCod, , oCon.ConexionActiva, TCF)
                                     CVHaber = CVHaber + nMovSoles
                                     
                                ElseIf Left(vCodConta, 4) = lsCajaDolares Then
                                     oMov.InsertaMovCta lnMovNro, lnItem, vCodConta, paAsiento * -1
                                     oMov.InsertaMovMe lnMovNro, lnItem, nImporteSoles * -1
                                     'CTI6 ERS112020
                                     'Call IngresarAsientoDN(FechaHora, vCodConta, 0, paAsiento, 1, vAgencia, pnMovNro, psOpeCod, , oCon.ConexionActiva, TCF)
                                     Call IngresarAsientoDN(FechaHora, vCodConta, 0, paAsiento, 1, vAgencia, pnMovNro, lsOpeCod, , oCon.ConexionActiva, TCF)
                                    
                                     Call IngresarAsientoDN(FechaHora, vCodConta, 0, nImporteSoles, 2, vAgencia, pnMovNro, lsOpeCod, , oCon.ConexionActiva, TCF)
                                     'Call IngresarAsientoDN(FechaHora, vCodConta, 0, nImporteSoles, 2, vAgencia, pnMovNro, psOpeCod, , oCon.ConexionActiva, TCF)
                                    CVHaber = CVHaber + paAsiento

                                End If
                            End If
                      RegOpeCta.MoveNext
                     Loop
               End If
            ' JUCS 20170822 -Ganancia y Perdida VC ME********************************************************************************************
              'si debe es menor al haber entonces se generarará una perdida
              vCtaDolar = CVDebe - CVHaber

             If vCtaDolar > 0 Then
                'CTI6 ERS112020
'                 sSql = "INSERT INTO AsientoDCVME (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge,nMovNro,cOpeCod) " & _
'                           " VALUES('" & FechaHora & "','" & lsCVMEGanacia & vAgencia & "',0," & Abs(vCtaDolar) & ",'1','" & vAgencia & "','" & pnMovNro & "','" & psOpeCod & "')"
                 sSql = "INSERT INTO AsientoDCVME (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge,nMovNro,cOpeCod) " & _
                           " VALUES('" & FechaHora & "','" & lsCVMEGanacia & vAgencia & "',0," & Abs(vCtaDolar) & ",'1','" & vAgencia & "','" & pnMovNro & "','" & lsOpeCod & "')"
                  oCon.Ejecutar sSql
                        lnItem = lnItem + 1
                        'vCtaDolar = Abs(vCtaDolar)
                        oMov.InsertaMovCta lnMovNro, lnItem, lsCVMEGanacia & vAgencia, vCtaDolar * -1
              ElseIf vCtaDolar < 0 Then
              'CTI6 ERS112020
'                  sSql = "INSERT INTO AsientoDCVME (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge,nMovNro,cOpeCod) " & _
'                            " VALUES('" & FechaHora & "','" & lsCVMEPerdida & vAgencia & "'," & Abs(vCtaDolar) & ",0,'1','" & vAgencia & "','" & pnMovNro & "','" & psOpeCod & "')"
                sSql = "INSERT INTO AsientoDCVME (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge,nMovNro,cOpeCod) " & _
                            " VALUES('" & FechaHora & "','" & lsCVMEPerdida & vAgencia & "'," & Abs(vCtaDolar) & ",0,'1','" & vAgencia & "','" & pnMovNro & "','" & lsOpeCod & "')"

                  oCon.Ejecutar sSql
                     lnItem = lnItem + 1
                     oMov.InsertaMovCta lnMovNro, lnItem, lsCVMEPerdida & vAgencia, Abs(vCtaDolar)
            End If
            'CTI6 ERS0112020 gCVTipoPagoEfectivo
            If pnTipoPago = gCVTipoPagoCargoCta Then
                Dim oCapta As COMDCaptaGenerales.DCOMCaptaMovimiento
                Set oCapta = New COMDCaptaGenerales.DCOMCaptaMovimiento
    
                Dim lsFechaHoraGrab As String
                lsFechaHoraGrab = SacarFechaHoraGrab(psMovNro)
                oMov.CapCargoCuentaAho pMatAho, psCtaCodCargo, pnMontoCargoCta + pnMontoITF, lsOpeCodCargoCta, psMovNro, lsOpeCodCargoCtaDescripcion, lsFechaHoraGrab, , , , , , , , , , , , , , , , , , , pnTipoPago
            End If
            
             If pnTipoPago = gCVTipoPagoVoucher Then
                Dim nMoneda As Integer
                nMoneda = pnMonedaTrans
                Dim oGen As New COMDConstSistema.DCOMGeneral
                If pnMovNroTransfer <> -1 Then oMov.AgregaMovRef lnMovNro, pnMovNroTransfer
                If pnMovNroTransfer <> -1 Then oMov.ActualizaMovPendientesRend pnMovNroTransfer, pnImporte 'pnMontoRVD
                oMov.actualizarRelacionVoucherCaptacion lnMovNro, pnMovNroRVD
             End If
              
            'END CTI6
            'Aqui ira lo relacionado a ajuste de saldo
            oMov.ActualizaSaldoMovimiento psMovNro, "+"
             '***************************************************************************************************************************************************
            'Fin TI-ERS 002-2017 20170822****************************************************************************************************************
                    RegOpeCta.Close
               Set RegOpeCta = Nothing
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
'CTI6 RS0112020
Public Function SacarFechaHoraGrab(ByVal psMovNro As String) As String
    SacarFechaHoraGrab = Mid(psMovNro, 1, 4) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 7, 2) & " " & Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2)
End Function
'JUCS 20170822, Segun TI-ERS 002-2017-obtenido de gFunAsientoD para usarlo en la obtencion de plantilla de la Ope CVME
Public Function VarAG(ByVal pCuentaProd As String, ByVal pCtaCont As String, Optional ByVal oCon As adodb.Connection = Nothing) As String
Dim lsAge As String
    '******** ica parche en caja agencias EJRS 13/10/2004 *****************************
    If Left(pCtaCont, 2) = "11" Then
        lsAge = Mid(pCuentaProd, 4, 2)
    Else
        lsAge = Mid(pCuentaProd, 4, 2)
        '*** PEAC 20140204
        Select Case lsAge
            Case "08"
                Dim rs As adodb.Recordset
                Dim sql As String
                Set rs = New adodb.Recordset
                sql = "exec stp_sel_BuscaAgeColocaciones '" & pCuentaProd & "' "
                Set rs = oCon.Execute(sql)
                If Not rs.EOF And Not rs.BOF Then
                    lsAge = Trim(rs!cAgeCodAct)
                End If
                rs.Close
                Set rs = Nothing
        End Select
        '*** FIN PEAC
    End If
    '**** CULMINACION DE PARCHES JEAN 13/10/2004 ***********+
    VarAG = lsAge
End Function
'JUCS 20170822, Segun TI-ERS 002-2017-obtenido de gFunAsientoD para usarlo en la obtencion de plantilla de la Ope CVME
Public Function AsientoParche(pAsiento As String, Optional ByVal pNuePlan As Boolean = False, Optional oCon As adodb.Connection) As String
Dim tmpRs As New adodb.Recordset
    Dim tmpSql As String
    If pNuePlan Then
        tmpSql = " SELECT cCodCntNue" & _
            " FROM AsientoPuenteN " & _
            " WHERE cCodCnt = '" & Trim(pAsiento) & "'"
    Else
        tmpSql = " SELECT cCodCntNue" & _
            " FROM AsientoPuente " & _
            " WHERE cCodCnt = '" & Trim(pAsiento) & "'"
    End If
    Set tmpRs = New adodb.Recordset
    'ALPA***ASIENTO
    Set tmpRs = oCon.Execute(tmpSql)
    If (tmpRs.BOF Or tmpRs.EOF) Then
        AsientoParche = ""
    Else
        AsientoParche = Trim(tmpRs!cCodCntNue)
    End If
    tmpRs.Close
    Set tmpRs = Nothing
    '***ALPA***20080604****************
End Function
'JUCS 20170822, Segun TI-ERS 002-2017-obtenido de gFunAsientoD para usarlo en la obtencion de plantilla de la Ope CVME
Public Sub IngresarAsientoDN(pdHoraGrab As String, psCodConta As String, pnMontoOperacionD As Currency, pnMontoOperacionH As Currency, psTipo As String, psCodAge As String, Optional ByVal pnMovNro As Long = -1, Optional psOpeCod As String = "", Optional psCtaCod As String = "", Optional oCon As adodb.Connection = Nothing, Optional ByVal pnTipoCambio As Currency = 0#)
Dim sSql As String
'Dim psFecha As String
Dim lbInactivas As Boolean
  'aqui ingresa para insertar en AsientoDN
   If (lbInactivas = 0 And psOpeCod = "200802") Or psOpeCod <> "200802" Then
         sSql = "INSERT INTO AsientoDCVME (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) "
         sSql = sSql & " VALUES('" & pdHoraGrab & "','" & psCodConta & "'," & pnMontoOperacionD & ", " & pnMontoOperacionH & ", '" & psTipo & "' ,'" & psCodAge & "'" 'Modificado PASI20140430
         'Agregado PASI20140430
         If pnMovNro = -1 Then
              sSql = sSql & ", Null "
         Else
              sSql = sSql & "," & pnMovNro
         End If
         If psOpeCod = "" Then
              sSql = sSql & ", Null "
         Else
              sSql = sSql & ",'" & psOpeCod & "'"
         End If
         If psCtaCod = "" Then
              sSql = sSql & ", Null "
         Else
              sSql = sSql & ",'" & psCtaCod & "'"
         End If
         sSql = sSql & ")"
         'END PASI
        oCon.Execute sSql
  End If
End Sub
'FIN JUCS

'ANDE 20170515 DINERO ELECTRÓNICO
Public Function GrabaConversionReconversionDE(ByVal pcMSISDN As String, ByVal pcDNI As String, ByVal pcPersNombre As String, _
                                                ByVal pcPersDomicilio As String, _
                                                ByVal pcPersCod As String, _
                                                ByVal pnMonto As Currency, ByVal pcPrdConcepto As String, _
                                                ByVal pcConceptoDesc As String, ByVal pcOpeCod As CaptacOperacion, _
                                                ByVal pnTipoMov As Integer, ByVal pdFecha As Date, _
                                                ByVal pcCodAge As String, ByVal pcCodUser As String, _
                                                Optional ByRef pcMovNro As String, _
                                                Optional ByVal pnTipoPago As Integer = 1, Optional ByVal psCtaCodCargo As String = "", _
                                                Optional ByRef pMatAho As Variant = "", _
                                                Optional ByVal pbITFAplica As Boolean, _
                                                Optional ByVal pnMontoITF As Currency, _
                                                Optional ByVal pnMontoCargoCta As Currency = 0#, Optional ByVal psNomAge As String, _
                                                Optional ByRef psImpBoleta As String, Optional pbImpTMU As Boolean, _
                                                Optional ByVal psLpt As String = "") As Integer

    Dim clsCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsCont As COMNContabilidad.NCOMContFunciones
    Dim cMov As String
    Dim nMovNro As Long
    
    On Error GoTo RegistrarConversionReConversion ' try
    
    Set clsCont = New COMNContabilidad.NCOMContFunciones
    Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento

    cMov = clsCont.GeneraMovNro(pdFecha, pcCodAge, pcCodUser)
    pcMovNro = cMov
    
    'CTI7 OPEv2
    Select Case pnTipoPago
        Case gColocTipoPagoEfectivo   ' Efectivo
                pcOpeCod = pcOpeCod
          Case gColocTipoPagoCargoCta
                pcOpeCod = gOpeTransferenciaCargo
          Case gColocTipoPagoDeposito
                pcOpeCod = gOpeTransferenciaDeposito
    End Select
    'End CTI7
    Dim lsFechaHoraGrab As String
    lsFechaHoraGrab = SacarFechaHoraGrab(pcMovNro)
    nMovNro = clsCapMov.OtrasOperaciones(cMov, pcOpeCod, pnMonto, pcMSISDN & "-" & CStr(pdFecha), pcConceptoDesc, "1", pcPersCod, , , , , , , , , , pcPrdConcepto, psCtaCodCargo, pMatAho, pnTipoPago, lsFechaHoraGrab, , , , , psNomAge, , , , , , psImpBoleta, pbImpTMU, psLpt, pcPersNombre)
    
    Dim sSql As String
    Dim oConect As COMConecta.DCOMConecta
    Set oConect = New COMConecta.DCOMConecta
    sSql = "stp_ins_DE_RegistrarConversionREConversion '" & CStr(nMovNro) & "','" & pcPersCod & "','" & pnTipoMov & "','" & CStr(pnMonto) & "','" & pcDNI & "','" & pcPersNombre & "','" & pcPersDomicilio & "','" & pcMSISDN & "'"
    oConect.AbreConexion
    If (nMovNro > 0) Then
        oConect.Ejecutar sSql
    End If
    oConect.CierraConexion
    GrabaConversionReconversionDE = 0
    Exit Function
    'catch (exception)
RegistrarConversionReConversion:
    Err.Raise Err.Number, "Error al tratar de guardar los datos de entrega de premio.", Err.Description
End Function
'FIN ANDE


Public Function GetCompraVenta(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                              ByVal psCodUser As String) As adodb.Recordset
Dim sql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset
Dim lsFiltroUser As String

Set rs = New adodb.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
'JUEZ 20150922 ***********************************************
'If psCodUser <> "" Then
'    lsFiltroUser = "  AND RIGHT(M.CMOVNRO,4)='" & psCodUser & "' "
'End If
'
'Sql = " Select CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) AS FECHA, " _
'    & "         O.COPEDESC, RIGHT(M.CMOVNRO,4) AS CUSER, P.CPERSNOMBRE, cSimbolo = 'US$', MC.NMOVIMPORTE, " _
'    & "         M.nMovNro, nMoneda = 2,  M.cMovDesc " _
'    & " From    MOV M " _
'    & "         JOIN MovCompraVenta MC  ON MC.NMOVNRO =M.NMOVNRO " _
'    & "         JOIN MOVTPOCAMBIO MT    ON MT.NMOVNRO = MC.NMOVNRO " _
'    & "         JOIN " & vsBasePesonas & "PERSONA P      ON P.CPERSCOD = MC.CPERSCOD " _
'    & "         JOIN " & vsBaseComunes & "OPETPO O       ON O.COPECOD = M.COPECOD " _
'    & " WHERE   M.NMOVFLAG =" & gMovFlagVigente & " AND M.COPECOD='" & psOpeCod & "' " & lsFiltroUser _
'    & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' and  '" & Format(pdHasta, "yyyymmdd") & "' " _
'    & " ORDER BY M.NMOVNRO "
sql = "stp_sel_GetCompraVenta '" & psOpeCod & "','" & Format(pdDesde, "yyyyMMdd") & "','" & Format(pdHasta, "yyyyMMdd") & "','" & psCodUser & "'"
'END JUEZ ****************************************************
 
Set rs = oConect.CargaRecordSet(sql)
Set GetCompraVenta = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function ValidaReciboServ(ByVal psOpeCod As String, ByVal psNroRecibo As String, ByVal psReferencia As String) As Boolean
Dim sql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset
Dim lsFiltroUser As String

Set rs = New adodb.Recordset
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
Dim oMov As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New COMDMov.DCOMMov

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
                              ByVal psCodUser As String, ByVal psCodAge As String) As adodb.Recordset
Dim sSql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset
Dim lsFiltroUser As String

Set rs = New adodb.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
If psCodUser <> "" Then
    lsFiltroUser = "  AND RIGHT(M.cMovNro,4)='" & psCodUser & "' "
End If

sSql = "SELECT CONVERT(Varchar(10),CONVERT(DATETIME, SUBSTRING(M.cMovNro,1,8)),103) + ' ' + " _
    & "SUBSTRING(M.cMovNro,9,2) + ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) Fecha, " _
    & "O.cOpeDesc, RIGHT(M.cMovNro,4) Usu, P.cPersNombre, cSimbolo = 'S/.', MS.nMonto, M.nMovNro, " _
    & "nMoneda = 1, M.cMovDesc FROM OpeTpo O JOIN Mov M JOIN MovServicios MS ON MS.nMovNro = M.nMovNro " _
    & "ON O.cOpeCod = M.cOpeCod JOIN RRHH RH JOIN Persona P ON RH.cPersCod = P.cPersCod ON " _
    & "RIGHT(M.cMovNro,4) = RH.cUser WHERE M.nMovFlag = " & gMovFlagVigente & " AND " _
    & "M.cOpeCod = '" & psOpeCod & "' And SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " & lsFiltroUser _
    & "AND SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' And '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "ORDER BY M.nMovNro "
 
Set rs = oConect.CargaRecordSet(sSql)
Set GetServicios = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetMovFideicomiso(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                              ByVal psCodUser As String, ByVal psCodAge As String) As adodb.Recordset
Dim sSql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset
Dim lsFiltroUser As String

Set rs = New adodb.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
If psCodUser <> "" Then
    lsFiltroUser = "  AND RIGHT(M.cMovNro,4)='" & psCodUser & "' "
End If

sSql = "SELECT CONVERT(Varchar(10),CONVERT(DATETIME, SUBSTRING(M.cMovNro,1,8)),103) + ' ' + " _
    & "SUBSTRING(M.cMovNro,9,2) + ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) Fecha, " _
    & "O.cOpeDesc, RIGHT(M.cMovNro,4) Usu, F.cCtaCod, cSimbolo = 'S/.', MS.nMonto, M.nMovNro, " _
    & "nMoneda = 1, M.cMovDesc FROM OpeTpo O JOIN Mov M JOIN MovServicios MS ON MS.nMovNro = M.nMovNro " _
    & "ON O.cOpeCod = M.cOpeCod INNER JOIN FIDEICOMISOMOV FM ON FM.nNroMov = MS.nMovNro " _
    & " INNER JOIN FIDEICOMISO F ON FM.cCtaCod = F.cCtaCod WHERE M.nMovFlag = " & gMovFlagVigente & " AND " _
    & "M.cOpeCod = '" & psOpeCod & "' And SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " & lsFiltroUser _
    & "AND SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' And '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "ORDER BY M.nMovNro "
 
Set rs = oConect.CargaRecordSet(sSql)
Set GetMovFideicomiso = rs
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetIngEfectFalt(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
            ByVal psCodUser As String, ByVal psCodAge As String) As adodb.Recordset
Dim sSql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset
Dim lsFiltroUser As String

Set rs = New adodb.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.cMovNro,4) = '" & psCodUser & "' "
End If

sSql = "SELECT CONVERT(CHAR(10),CONVERT(DATETIME, SUBSTRING(M.cMovNro,1,8)),103) + ' ' + SUBSTRING(M.cMovNro,9,2)+ ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) AS  FECHA, " _
    & "O.cOpeDesc, RIGHT(M.cMovNro,4) AS cUser, CONVERT(CHAR(25),MS.cNroDoc) + MS.cReferencia AS cDocumentos, " _
    & "cSimbolo = CASE WHEN MS.nMoneda = " & gMonedaNacional & " THEN 'S/.' ELSE 'US$' END, MS.nMovImporte, M.nMovNro, MS.nMoneda, M.cMovDesc FROM MOV M JOIN MovOpeVarias MS ON MS.nMovNro = M.nMovNro " _
    & "JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod WHERE M.nMovFlag = " & gMovFlagVigente & " AND " _
    & "M.cOpeCod In ('" & psOpeCod & "') AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " & lsFiltroUser _
    & "AND SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format$(pdDesde, "yyyymmdd") & "' And '" & Format$(pdHasta, "yyyymmdd") & "' " _
    & "AND EXISTS (SELECT  M1.nMovNro FROM MOV M1 JOIN MOVREF MR ON MR.nMovNroREF = M1.nMovNro " _
    & "WHERE M1.nMovFlag = " & gMovFlagVigente & " AND MR.nMovNro = M.nMovNro) " _
    & "ORDER BY M.nMovNro "
 
Set rs = oConect.CargaRecordSet(sSql)
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
                                        Optional ByVal psCodCmac As String) As adodb.Recordset
    Dim sSql As String
    Dim rs As adodb.Recordset
    Dim lsFiltroUser As String
    Dim lsFiltroAge As String
    Dim oConect As COMConecta.DCOMConecta
    Dim lnTipoGrupo As GruposIngEgre
    
    
    Dim lsTablaDiaria As String
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    
    'If pnTipoGrupo = gGruposIngEgreOtraAgencia Then
    '    lnTipoGrupo = gGruposIngEgreAgeLocal
    'Else
    '    lnTipoGrupo = pnTipoGrupo
    'End If
    '
    'If pdFecha = pdFechaDia Then
    '    lsTablaDiaria = " MovDiario "
    'Else
    '    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
    '    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
    'End If
    
    Set rs = New adodb.Recordset
    ' If psCodUser <> "" Then
    '     lsFiltroUser = " AND RIGHT(M.cMovNro,4)  = '" & psCodUser & "' "
    ' End If
    '
    '
    ' If psCodCmac = "102" Then
    '     'Solo para Lima
    '     sSql = "SELECT cGrupoNombre, SUM(nTotalMov) AS nTotalMov, " _
    '         & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
    '         & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END ) AS Cheque , " _
    '         & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, G.cGrupoCod " _
    '         & "FROM GrupoOpe G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
    '         & "JOIN (SELECT Mc.cOpeCod, O.cOpeDesc, COUNT(nmovnro) AS nTotalMov, " _
    '         & "  CASE WHEN MD.nDocTpo IS NULL THEN SUM(MC.nMonto) ELSE 0 END AS Efectivo , " _
    '         & "  CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN  SUM(MC.nMonto) ELSE 0 END AS Cheque, " _
    '         & "  CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN SUM(MC.nMonto) ELSE 0 END AS OrdenPago " _
    '         & "  FROM Mov M JOIN MovCap MC ON MC.nMovNro = M.nMovNro  " _
    '         & "  JOIN (SELECT NMOVNRO,COPECOD,CCTACOD,NCONCEPTOCOD FROM MOVCAPDET WHERE NCONCEPTOCOD NOT IN (21,3)) MCD ON   MCD.NMOVNRO=MC.NMOVNRO AND MCD.CCTACOD=MC.CCTACOD AND MCD.COPECOD=MC.COPECOD " _
    '         & "  LEFT JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
    '         & "  JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
    '         & "  WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
    '         & "  AND SUBSTRING(MC.cCtaCod,9,1)='" & pnMoneda & "' And M.nMovFlag =" & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
    '         & "  GROUP BY MC.cOpeCod, O.cOpeDesc, MD.nDocTpo) As OpeCol  " _
    '         & "ON OpeCol.cOpeCod = GO.cOpeCod " _
    '         & "WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%' " _
    '         & "GROUP BY cGrupoNombre, G.cGrupoCod "
    ' Else
    '     sSql = " SELECT CASE WHEN OpeCol.cCodAge = '" & psCodAge & "' THEN '000' ELSE (OpeCol.cCodAge + '2') END as nAgeOrd, OpeCol.cCodAge, cGrupoNombre, SUM(nTotalMov) AS nTotalMov, " _
    '         & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
    '         & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END ) AS Cheque , " _
    '         & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, G.cGrupoCod " _
    '         & "FROM GrupoOpe G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
    '         & "JOIN (SELECT SUBSTRING(MC.cCtaCod,4,2) AS cCodAge, Mc.cOpeCod, O.cOpeDesc, COUNT(DISTINCT m.nmovnro) AS nTotalMov, " _
    '         & " CASE WHEN (MD.nDocTpo IS NULL OR MC.COPECOD LIKE '99%') THEN SUM(MC.nMonto) ELSE 0 END AS Efectivo , " _
    '         & " CASE WHEN (MD.nDocTpo = " & TpoDocCheque & " And MC.cOpeCod NOT LIKE '99%') THEN  SUM(MC.nMonto) ELSE 0 END AS Cheque, " _
    '         & " CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " AND NOT MC.COPECOD LIKE '99%' THEN SUM(MC.nMonto) ELSE 0 END AS OrdenPago " _
    '         & " FROM " & lsTablaDiaria & " M JOIN MovCap MC ON MC.nMovNro = M.nMovNro JOIN " _
    '         & " (Select nMovNro, cOpeCod, cCtaCod From MovCapDet Where nConceptoCod NOT IN (21,23) Group by nMovNro, cOpeCod, cCtaCod) MCD ON MC.nMovNro = MCD.nMovNro And MC.cCtaCod = MCD.cCtaCod And MC.cOpeCod = MCD.cOpeCod " _
    '         & " LEFT JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
    '         & " JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
    '         & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
    '         & " AND SUBSTRING(MC.cCtaCod,9,1)='" & pnMoneda & "' And M.nMovFlag =" & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
    '         & " GROUP BY Mc.cOpeCod, O.cOpeDesc, MD.nDocTpo, SUBSTRING(MC.cCtaCod,4,2) "
    '
    '      sSql = sSql & "UNION" _
    '         & " SELECT SUBSTRING(M.cMovNro,18,2) AS cCodAge, M.cOpeCod, O.cOpeDesc, COUNT(m.nmovnro) AS nTotalMov," _
    '         & " CASE WHEN MD.nDocTpo IS NULL THEN SUM(MV.nMovImporte) ELSE 0 END AS Efectivo , " _
    '         & " CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN SUM(MV.nMovImporte) ELSE 0 END AS Cheque," _
    '         & " CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN SUM(MV.nMovImporte) ELSE 0 END AS OrdenPago " _
    '         & " FROM " & lsTablaDiaria & " M " _
    '         & " JOIN MovOpevarias MV ON MV.nMovNro = M.nMovNro" _
    '         & " JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ")" _
    '         & " JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod" _
    '         & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "'" & lsFiltroUser & " " _
    '         & " AND MV.nMoneda = '" & pnMoneda & "' And M.nMovFlag = 0 AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
    '         & " GROUP BY M.cOpeCod, O.cOpeDesc, MD.nDocTpo, SUBSTRING(M.cMovNro,18,2)"
    '
    '     sSql = sSql & ") As OpeCol  " _
    '         & " ON OpeCol.cOpeCod = GO.cOpeCod " _
    '         & " WHERE G.cGrupoCod LIKE '" & Format(lnTipoGrupo, "00") & "%' " _
    '         & " GROUP BY OpeCol.cCodAge, cGrupoNombre, G.cGrupoCod  Having " & IIf(pnTipoGrupo = gGruposIngEgreAgeLocal, " OpeCol.cCodAge = '" & psCodAge & "'", " OpeCol.cCodAge <> '" & psCodAge & "'")
    '     sSql = sSql + " UNION " _
    '         & " (SELECT DISTINCT CASE WHEN cAgeCod = '" & psCodAge & "' THEN '000' ELSE  (cAgeCod + '1') END as nAgeOrd, cAgeCod, cAgeDescripcion,0,0,0,0,'' " _
    '         & " FROM Agencias AG JOIN " _
    '         & " (SELECT SUBSTRING(MC.cCtaCod,4,2) As cCodAge " _
    '         & "  FROM MovCap MC JOIN " & lsTablaDiaria & " M ON M.nMovNro = MC.nMovNro " _
    '         & "   JOIN (SELECT NMOVNRO,COPECOD,CCTACOD,NCONCEPTOCOD FROM MOVCAPDET WHERE NCONCEPTOCOD<>21) MCD ON   MCD.NMOVNRO=MC.NMOVNRO AND MCD.CCTACOD=MC.CCTACOD AND MCD.COPECOD=MC.COPECOD " _
    '         & "  WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
    '         & "  AND SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
    '         & " ) Mov ON Mov.cCodAge = Ag.cAgeCod " & IIf(pnTipoGrupo = gGruposIngEgreAgeLocal, " And cAgeCod = '" & psCodAge & "'", " And cAgeCod <> '" & psCodAge & "' ") & " ) " _
    '         & "  ORDER BY nAgeOrd "
    ' End If '
    
    sSql = "sp_sel_GetOpeIngEgreCaptaciones " & pnTipoGrupo & ", '" & Format(pdFecha, "yyyymmdd") & "', '" _
                                              & psCodAge & "', '" & psCodUser & "', " & pnMoneda & ", '" _
                                              & Format(pdFechaDia, "yyyymmdd") & "', '" & pgsCodUser & "',0"
    
    Set rs = oConect.CargaRecordSet(sSql)
    Set GetOpeIngEgreCaptaciones = rs
    'If pdFecha <> pdFechaDia Then
    '    Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
    'End If
    oConect.CierraConexion: Set oConect = Nothing
    End Function

'MADM 20100130
Public Function GetOpeIngEgreCaptacionesPreCuadre(ByVal pnTipoGrupo As GruposIngEgre, _
                                                  ByVal pdFecha As Date, _
                                                  ByVal psCodAge As String, _
                                                  ByVal psCodUser As String, _
                                                  ByVal pnMoneda As Moneda, _
                                                  pdFechaDia As Date, _
                                                  pgsCodUser As String, _
                                                  Optional ByVal psCodCmac As String, _
                                                  Optional ByVal pbind As Boolean = False) As adodb.Recordset
    Dim sSql As String
    Dim rs As adodb.Recordset
    Dim lsFiltroUser As String
    Dim lsFiltroAge As String
    Dim oConect As COMConecta.DCOMConecta
    Dim lnTipoGrupo As GruposIngEgre
    Dim lsFiltroPre As String
    
    Dim lsTablaDiaria As String
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    
    'RIRO20151229 COMENTADO
    'If pnTipoGrupo = gGruposIngEgreOtraAgencia Then
    '    lnTipoGrupo = gGruposIngEgreAgeLocal
    'Else
    '    lnTipoGrupo = pnTipoGrupo
    'End If
    '
    'If pdFecha = pdFechaDia Then
    '    lsTablaDiaria = " MovDiario "
    'Else
    '    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
    '    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
    'End If
    
    Set rs = New adodb.Recordset
    
    ' If psCodUser <> "" Then
    '     lsFiltroUser = " AND RIGHT(M.cMovNro,4)  = '" & psCodUser & "' "
    ' End If
    '
    'If pbind Then
    '  lsFiltroPre = " and M.nMovNro > (Select max(E.nMovNro) nMovNro From  MovDiario  M INNER JOIN MOv E ON M.nMovNro = E.nMovNro  "
    '  lsFiltroPre = lsFiltroPre & " Where SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' AND M.nMovFlag NOT IN (2,1) and SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "'" & lsFiltroUser & " AND M.cOpeCod in ('901018')  ) "
    'End If
    '
    'sSql = " SELECT CASE WHEN OpeCol.cCodAge = '" & psCodAge & "' THEN '000' ELSE (OpeCol.cCodAge + '2') END as nAgeOrd, OpeCol.cCodAge, cGrupoNombre, SUM(nTotalMov) AS nTotalMov, " _
    '         & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
    '         & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END ) AS Cheque , " _
    '         & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, G.cGrupoCod " _
    '         & "FROM GrupoOpe G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
    '         & "JOIN (SELECT SUBSTRING(MC.cCtaCod,4,2) AS cCodAge, Mc.cOpeCod, O.cOpeDesc, COUNT(DISTINCT m.nmovnro) AS nTotalMov, " _
    '         & " CASE WHEN (MD.nDocTpo IS NULL OR MC.COPECOD LIKE '99%') THEN SUM(MC.nMonto) ELSE 0 END AS Efectivo , " _
    '         & " CASE WHEN (MD.nDocTpo = " & TpoDocCheque & " And MC.cOpeCod NOT LIKE '99%') THEN  SUM(MC.nMonto) ELSE 0 END AS Cheque, " _
    '         & " CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " AND NOT MC.COPECOD LIKE '99%' THEN SUM(MC.nMonto) ELSE 0 END AS OrdenPago " _
    '         & " FROM " & lsTablaDiaria & " M JOIN MovCap MC ON MC.nMovNro = M.nMovNro JOIN " _
    '         & " (Select nMovNro, cOpeCod, cCtaCod From MovCapDet Where nConceptoCod NOT IN (21) " _
    '         & " Group by nMovNro, cOpeCod, cCtaCod) MCD ON MC.nMovNro = MCD.nMovNro And MC.cCtaCod = MCD.cCtaCod And MC.cOpeCod = MCD.cOpeCod " & lsFiltroPre _
    '         & " LEFT JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
    '         & " JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
    '         & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
    '         & " AND SUBSTRING(MC.cCtaCod,9,1)='" & pnMoneda & "' And M.nMovFlag =" & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
    '         & " GROUP BY Mc.cOpeCod, O.cOpeDesc, MD.nDocTpo, SUBSTRING(MC.cCtaCod,4,2) "
    '
    '      sSql = sSql & "UNION" _
    '         & " SELECT SUBSTRING(M.cMovNro,18,2) AS cCodAge, M.cOpeCod, O.cOpeDesc, COUNT(m.nmovnro) AS nTotalMov," _
    '         & " CASE WHEN MD.nDocTpo IS NULL THEN SUM(MV.nMovImporte) ELSE 0 END AS Efectivo , " _
    '         & " CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN SUM(MV.nMovImporte) ELSE 0 END AS Cheque," _
    '         & " CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN SUM(MV.nMovImporte) ELSE 0 END AS OrdenPago " _
    '         & " FROM " & lsTablaDiaria & " M " _
    '         & " JOIN MovOpevarias MV ON MV.nMovNro = M.nMovNro" _
    '         & " JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ")" _
    '         & " JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod" _
    '         & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "'" & lsFiltroUser & " " _
    '         & " AND MV.nMoneda = '" & pnMoneda & "' And M.nMovFlag = 0 AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
    '         & " GROUP BY M.cOpeCod, O.cOpeDesc, MD.nDocTpo, SUBSTRING(M.cMovNro,18,2)"
    '
    '     sSql = sSql & ") As OpeCol  " _
    '         & " ON OpeCol.cOpeCod = GO.cOpeCod " _
    '         & " WHERE G.cGrupoCod LIKE '" & Format(lnTipoGrupo, "00") & "%' " _
    '         & " GROUP BY OpeCol.cCodAge, cGrupoNombre, G.cGrupoCod  Having " & IIf(pnTipoGrupo = gGruposIngEgreAgeLocal, " OpeCol.cCodAge = '" & psCodAge & "'", " OpeCol.cCodAge <> '" & psCodAge & "'")
    '     sSql = sSql + " UNION " _
    '         & " (SELECT DISTINCT CASE WHEN cAgeCod = '" & psCodAge & "' THEN '000' ELSE  (cAgeCod + '1') END as nAgeOrd, cAgeCod, cAgeDescripcion,0,0,0,0,'' " _
    '         & " FROM Agencias AG JOIN " _
    '         & " (SELECT SUBSTRING(MC.cCtaCod,4,2) As cCodAge " _
    '         & "  FROM MovCap MC JOIN " & lsTablaDiaria & " M ON M.nMovNro = MC.nMovNro " _
    '         & "   JOIN (SELECT NMOVNRO,COPECOD,CCTACOD,NCONCEPTOCOD FROM MOVCAPDET WHERE NCONCEPTOCOD<>21) MCD ON   MCD.NMOVNRO=MC.NMOVNRO AND MCD.CCTACOD=MC.CCTACOD AND MCD.COPECOD=MC.COPECOD " _
    '         & "  WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
    '         & "  AND SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
    '         & " ) Mov ON Mov.cCodAge = Ag.cAgeCod " & IIf(pnTipoGrupo = gGruposIngEgreAgeLocal, " And cAgeCod = '" & psCodAge & "'", " And cAgeCod <> '" & psCodAge & "' ") & " ) " _
    '         & "  ORDER BY nAgeOrd "
    
    'ADD RIRO20151229 ERS 142-2014
    sSql = "sp_sel_GetOpeIngEgreCaptaciones " & pnTipoGrupo & ", '" & Format(pdFecha, "yyyymmdd") & "', '" _
                                              & psCodAge & "', '" & psCodUser & "', " & pnMoneda & ", '" _
                                              & Format(pdFechaDia, "yyyymmdd") & "', '" & pgsCodUser & "', " & IIf(pbind, 1, 0)
    
    Set rs = oConect.CargaRecordSet(sSql)
    Set GetOpeIngEgreCaptacionesPreCuadre = rs
    'If pdFecha <> pdFechaDia Then
    '    Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
    'End If
    oConect.CierraConexion: Set oConect = Nothing
    End Function
    
Public Function GetOpeIngEgreColocacionesPreCuadre(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
            ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String, _
            Optional ByVal psCodCmac As String, Optional ByVal pbind As Boolean = False) As adodb.Recordset

Dim sSql As String
Dim rs As adodb.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim oConect As COMConecta.DCOMConecta
Dim lsTablaDiaria As String
Dim lnTipoGrupo  As GruposIngEgre
Dim lsFiltroPre As String

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function


'    If pnTipoGrupo = gGruposIngEgreOtraAgencia Then
'        lnTipoGrupo = gGruposIngEgreAgeLocal
'    Else
'        lnTipoGrupo = pnTipoGrupo
'    End If
'
'    If pbind Then
'     lsFiltroPre = " and M.nMovNro > (Select max(E.nMovNro) nMovNro From  MovDiario  M INNER JOIN MOv E ON M.nMovNro = E.nMovNro  "
'     lsFiltroPre = lsFiltroPre & " Where SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' AND M.nMovFlag NOT IN (2,1) and SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "'" & lsFiltroUser & " AND M.cOpeCod in ('901018')  ) "
'   End If
'
'If pdFecha = pdFechaDia Then
'    lsTablaDiaria = " MovDiario "
'Else
'    'psCodUser = IIf(psCodUser = "", "9999", psCodUser)
'    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
'    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
'End If


Set rs = New adodb.Recordset

'If psCodUser <> "" Then
'    lsFiltroUser = " AND RIGHT(M.cMovNro,4)  = '" & psCodUser & "' "
'End If
'
'    sSql = "SELECT CASE WHEN OpeCol.cCodAge = '" & psCodAge & "' THEN '000' ELSE (OpeCol.cCodAge + '2') END as nAgeOrd, OpeCol.cCodAge, cGrupoNombre, SUM(nTotalMov) AS nTotalMov, " _
'        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
'        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END ) AS Cheque, " _
'        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, G.cGrupoCod " _
'        & "FROM GrupoOpe G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
'        & "JOIN (SELECT SUBSTRING(MC.cCtaCod,4,2) AS cCodAge, MC.cOpeCod, O.cOpeDesc, COUNT(M.NMOVNRO) AS nTotalMov, " _
'        & "  CASE WHEN MD.nDocTpo IS NULL THEN SUM(MC.nMonto) ELSE 0 END AS Efectivo, " _
'        & "  CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN  SUM(MC.nMonto) ELSE 0 END AS Cheque, " _
'        & "  CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN SUM(MC.nMonto) ELSE 0 END AS OrdenPago " _
'        & "  FROM " & lsTablaDiaria & " M JOIN MovCol MC ON MC.nMovNro = M.nMovNro " _
'        & "  LEFT JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
'        & "  JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
'        & "  WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
'        & "  AND SUBSTRING(MC.cCtaCod,9,1)='" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " & lsFiltroPre _
'        & "  GROUP BY MC.cOpeCod, O.cOpeDesc, MD.nDocTpo, SUBSTRING(MC.cCtaCod,4,2) ) As OpeCol  " _
'        & "ON OpeCol.cOpeCod = GO.cOpeCod " _
'        & "WHERE G.cGrupoCod LIKE '" & Format(lnTipoGrupo, "00") & "%' " _
'        & "GROUP BY OpeCol.cCodAge, cGrupoNombre, G.cGrupoCod Having " & IIf(pnTipoGrupo = gGruposIngEgreAgeLocal, " OpeCol.cCodAge = '" & psCodAge & "'", " OpeCol.cCodAge <> '" & psCodAge & "'")
'    sSql = sSql + " UNION " _
'        & "(SELECT DISTINCT CASE WHEN cAgeCod = '" & psCodAge & "' THEN '000' ELSE  (cAgeCod + '1') END as nAgeOrd, cAgeCod, cAgeDescripcion,0,0,0,0,'' " _
'        & "FROM Agencias AG JOIN " _
'        & "(SELECT SUBSTRING(MC.cCtaCod,4,2) As cCodAge " _
'        & " FROM MovCol MC JOIN " & lsTablaDiaria & " M ON M.nMovNro = MC.nMovNro " _
'        & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser & lsFiltroAge _
'        & " AND SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And M.nMovFlag =" & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
'        & ") Mov ON Mov.cCodAge = Ag.cAgeCod   " & IIf(pnTipoGrupo = gGruposIngEgreAgeLocal, " And cAgeCod = '" & psCodAge & "'", " And cAgeCod <> '" & psCodAge & "' ") & ") " _
'        & " ORDER BY nAgeOrd "
    
    sSql = "sp_sel_GetOpeIngEgreColocaciones " & pnTipoGrupo & ", '" _
                                               & Format(pdFecha, "yyyymmdd") & "', '" _
                                               & psCodAge & "', '" _
                                               & psCodUser & "', " _
                                               & pnMoneda & ", '" _
                                               & Format(pdFechaDia, "yyyymmdd") & "', '" _
                                               & pgsCodUser & "', " & IIf(pbind, 1, 0)

Set rs = oConect.CargaRecordSet(sSql)
Set GetOpeIngEgreColocacionesPreCuadre = rs
'If pdFecha <> pdFechaDia Then
'        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
'End If
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetOpeIngEgreOpeCMACsPreCuadre(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, ByVal psCodAge As String, _
            ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String, Optional ByVal pbind As Boolean = False) As adodb.Recordset

Dim sSql As String
Dim rs As adodb.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim oConect As COMConecta.DCOMConecta
Set oConect = New COMConecta.DCOMConecta
Dim lsFiltroPre As String

If oConect.AbreConexion = False Then Exit Function

Dim lsTablaDiaria As String

If pdFecha = pdFechaDia Then
    lsTablaDiaria = " MovDiario "
Else
    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
End If
    
    If pbind Then
     lsFiltroPre = " and M.nMovNro > (Select max(E.nMovNro) nMovNro From  MovDiario  M INNER JOIN MOv E ON M.nMovNro = E.nMovNro  "
     lsFiltroPre = lsFiltroPre & " Where SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' AND M.nMovFlag NOT IN (2,1) and SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "'" & lsFiltroUser & " AND M.cOpeCod in ('901018')  ) "
   End If

Set rs = New adodb.Recordset
If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.cMovNro,4)  = '" & psCodUser & "' "
End If



sSql = "SELECT (OpeCol.cCodAge + '2') As nAgeOrd, OpeCol.cCodAge, cGrupoNombre, SUM(nTotalMov) AS nTotalMov, " _
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
    & "  AND MC.nMoneda = '" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " & lsFiltroPre _
    & "  GROUP BY M.cOpeCod, O.cOpeDesc, MD.nDocTpo, MC.cPersCod) As OpeCol  " _
    & "ON OpeCol.cOpeCod = GO.cOpeCod " _
    & "WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%' And G.nEfectivo = 1 " _
    & "GROUP BY OpeCol.cCodAge, cGrupoNombre, G.cGrupoCod "
sSql = sSql + " UNION " _
    & "(SELECT DISTINCT (P.cPersCod + '1') As nAgeOrd, P.cPersCod As cAgeCod, P.cPersNombre As cAgeDescripcion,0,0,0,0,'' " _
    & "FROM Persona P JOIN " _
    & "(SELECT MC.cPersCod As cCodAge " _
    & " FROM MovCMAC MC JOIN " & lsTablaDiaria & " M ON M.nMovNro = MC.nMovNro " _
    & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
    & " AND MC.nMoneda = '" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
    & ") Mov ON Mov.cCodAge = P.cPersCod) " _
    & " ORDER BY nAgeOrd "

Set rs = oConect.CargaRecordSet(sSql)
Set GetOpeIngEgreOpeCMACsPreCuadre = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetOpeIngEgreOtrasOpePreCuadre(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String, Optional ByVal pbind As Boolean = False) As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim lsFiltroUser2 As String
Dim lsFiltroAge2 As String
Dim oConect As COMConecta.DCOMConecta
Dim lsTablaDiaria As String
Dim lsFiltroPre As String

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

'RIRO 20200804 Comentado ***********

'    If pbind Then
'     lsFiltroPre = " and M.nMovNro > (Select max(E.nMovNro) nMovNro From  MovDiario  M INNER JOIN MOv E ON M.nMovNro = E.nMovNro  "
'     lsFiltroPre = lsFiltroPre & " Where SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' AND M.nMovFlag NOT IN (2,1) and SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "'" & lsFiltroUser & " AND M.cOpeCod in ('901018')  ) "
'   End If
'
'If pdFecha = pdFechaDia Then
'    lsTablaDiaria = " MovDiario "
'Else
'    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
'    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
'End If
'
'Set rs = New ADODB.Recordset
'If psCodAge <> "" Then
'    lsFiltroAge = " AND  SUBSTRING(M.CMOVNRO,18,2) = '" & psCodAge & "'"
'    lsFiltroAge2 = " AND  SUBSTRING(CMOVNRO,18,2) = '" & psCodAge & "'"
'End If
'If psCodUser <> "" Then
'    lsFiltroUser = " AND RIGHT(M.CMOVNRO,4) = '" & psCodUser & "' "
'    lsFiltroUser2 = " AND RIGHT(CMOVNRO,4) = '" & psCodUser & "' "
'End If
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
'    & "  FROM " & lsTablaDiaria & " M JOIN MOVOPEVARIAS MV ON MV.nMovNro = M.nMovNro " _
'    & "  LEFT JOIN MOVDOC MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
'    & "  JOIN OPETPO O ON O.cOpeCod = M.cOpeCod " _
'    & "  WHERE   SUBSTRING(M.CMOVNRO,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser & lsFiltroPre _
'    & "  AND M.NMOVFLAG =" & gMovFlagVigente & " AND MV.nMoneda =" & pnMoneda & " " _
'    & "  GROUP BY  M.cOpeCod, O.COPEDESC, MD.nDocTpo ) AS OPESERV " _
'    & "  ON OPESERV.cOpeCod = GO.cOpeCod " _
'    & " WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%'  " _
'    & " GROUP BY cGrupoNombre, GO.cGrupoCod "

'END RIRO 20200804 ******************

sql = "stp_sel_OtrasOpe '" & _
        Format(pdFecha, "yyyymmdd") & "', '" & _
        psCodAge & "', '" & psCodUser & "', " & _
        pnMoneda & ", '" & _
        Format(pdFechaDia, "yyyymmdd") & "', 1, " & _
        IIf(pbind, 1, 0)
    
Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreOtrasOpePreCuadre = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetOpeIngEgreServiciosPreCuadre(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String, Optional ByVal pbind As Boolean = False) As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim oConect As COMConecta.DCOMConecta
Dim lsFiltroPre As String
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

Dim lsTablaDiaria As String

    If pbind Then
     lsFiltroPre = " and M.nMovNro > (Select max(E.nMovNro) nMovNro From  MovDiario  M INNER JOIN MOv E ON M.nMovNro = E.nMovNro  "
     lsFiltroPre = lsFiltroPre & " Where SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' AND M.nMovFlag NOT IN (2,1) and SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "'" & lsFiltroUser & " AND M.cOpeCod in ('901018')  ) "
   End If

If pdFecha = pdFechaDia Then
    lsTablaDiaria = " MovDiario "
Else
    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
End If

Set rs = New adodb.Recordset
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
    & "  Cheque = 0, OrdenPago = 0 FROM " & lsTablaDiaria & " M JOIN MovServicios MV ON MV.nMovNro = M.nMovNro and MV.cOpecod <> '" & gITFGiroCancelEfect & "'" _
    & "  JOIN OPETPO O ON O.cOpeCod = M.cOpeCod " _
    & "  WHERE SUBSTRING(M.CMOVNRO,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser & lsFiltroPre _
    & "  AND M.NMOVFLAG = " & gMovFlagVigente & " And MV.nMoneda = " & pnMoneda & "" _
    & "  GROUP BY  M.cOpeCod, O.COPEDESC) AS OPESERV " _
    & "  ON OPESERV.cOpeCod = GO.cOpeCod " _
    & " WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%'  " _
    & " GROUP BY cGrupoNombre, GO.cGrupoCod "
    'NAGL 20190307 Agregó and MV.cOpecod <> '" & gITFGiroCancelEfect & "'"
Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreServiciosPreCuadre = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function

'END MADM

Public Function GetOpeIngEgreColocaciones(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
            ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String, _
            Optional ByVal psCodCmac As String) As adodb.Recordset

Dim sSql As String
Dim rs As adodb.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim oConect As COMConecta.DCOMConecta
Dim lsTablaDiaria As String
Dim lnTipoGrupo  As GruposIngEgre

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function


'If pnTipoGrupo = gGruposIngEgreOtraAgencia Then
'    lnTipoGrupo = gGruposIngEgreAgeLocal
'Else
'    lnTipoGrupo = pnTipoGrupo
'End If
'
'
'If pdFecha = pdFechaDia Then
'    lsTablaDiaria = " MovDiario "
'Else
'    'psCodUser = IIf(psCodUser = "", "9999", psCodUser)
'    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
'    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
'End If


Set rs = New adodb.Recordset

'If psCodUser <> "" Then
'    lsFiltroUser = " AND RIGHT(M.cMovNro,4)  = '" & psCodUser & "' "
'End If
'
'If psCodCmac = "102" Then
'    'Solo para Lima
'    sSql = "SELECT cGrupoNombre, SUM(nTotalMov) AS nTotalMov, " _
'        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
'        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END ) AS Cheque, " _
'        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, G.cGrupoCod " _
'        & "FROM GrupoOpe G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
'        & "JOIN (SELECT M.cOpeCod, O.cOpeDesc, COUNT(*) AS nTotalMov, " _
'        & "  CASE WHEN MD.nDocTpo IS NULL THEN SUM(MC.nMonto) ELSE 0 END AS Efectivo, " _
'        & "  CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN  SUM(MC.nMonto) ELSE 0 END AS Cheque, " _
'        & "  CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN SUM(MC.nMonto) ELSE 0 END AS OrdenPago " _
'        & "  FROM Mov M JOIN MovCol MC ON MC.nMovNro = M.nMovNro " _
'        & "  LEFT JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
'        & "  JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
'        & "  WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
'        & "  AND SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
'        & "  GROUP BY M.cOpeCod, O.cOpeDesc, MD.nDocTpo) As OpeCol  " _
'        & "ON OpeCol.cOpeCod = GO.cOpeCod " _
'        & "WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%' " _
'        & "GROUP BY cGrupoNombre, G.cGrupoCod "
'Else
'    sSql = "SELECT CASE WHEN OpeCol.cCodAge = '" & psCodAge & "' THEN '000' ELSE (OpeCol.cCodAge + '2') END as nAgeOrd, OpeCol.cCodAge, cGrupoNombre, SUM(nTotalMov) AS nTotalMov, " _
'        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
'        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END ) AS Cheque, " _
'        & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, G.cGrupoCod " _
'        & "FROM GrupoOpe G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
'        & "JOIN (SELECT SUBSTRING(MC.cCtaCod,4,2) AS cCodAge, MC.cOpeCod, O.cOpeDesc, COUNT(M.NMOVNRO) AS nTotalMov, " _
'        & "  CASE WHEN MD.nDocTpo IS NULL THEN SUM(MC.nMonto) ELSE 0 END AS Efectivo, " _
'        & "  CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN  SUM(MC.nMonto) ELSE 0 END AS Cheque, " _
'        & "  CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN SUM(MC.nMonto) ELSE 0 END AS OrdenPago " _
'        & "  FROM " & lsTablaDiaria & " M JOIN MovCol MC ON MC.nMovNro = M.nMovNro " _
'        & "  LEFT JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
'        & "  JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
'        & "  WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
'        & "  AND SUBSTRING(MC.cCtaCod,9,1)='" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
'        & "  GROUP BY MC.cOpeCod, O.cOpeDesc, MD.nDocTpo, SUBSTRING(MC.cCtaCod,4,2) "
'
'    'RIRO20140620 ERS017 ******
'    sSql = sSql & " UNION SELECT  max(SUBSTRING(M.cMovNro,18,2)) AS cCodAge, MC.cOpeCod, O.cOpeDesc, COUNT(M.NMOVNRO) AS nTotalMov,   " & _
'           " SUM(mc.nMovImporte) Efectivo, 0 Cheque, 0 OrdenPago FROM  MovDiario  M JOIN MovOpeVarias MC ON MC.nMovNro = M.nMovNro " & _
'           " JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser & _
'           " AND MC.nMoneda='" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " & _
'           " and mc.copecod = '300529' GROUP BY MC.cOpeCod, O.cOpeDesc"
'    'END RIRO *****************
'
'        sSql = sSql & ") As OpeCol  " _
'        & "ON OpeCol.cOpeCod = GO.cOpeCod " _
'        & "WHERE G.cGrupoCod LIKE '" & Format(lnTipoGrupo, "00") & "%' " _
'        & "GROUP BY OpeCol.cCodAge, cGrupoNombre, G.cGrupoCod Having " & IIf(pnTipoGrupo = gGruposIngEgreAgeLocal, " OpeCol.cCodAge = '" & psCodAge & "'", " OpeCol.cCodAge <> '" & psCodAge & "'")
'    sSql = sSql + " UNION " _
'        & "(SELECT DISTINCT CASE WHEN cAgeCod = '" & psCodAge & "' THEN '000' ELSE  (cAgeCod + '1') END as nAgeOrd, cAgeCod, cAgeDescripcion,0,0,0,0,'' " _
'        & "FROM Agencias AG JOIN " _
'        & "(SELECT SUBSTRING(MC.cCtaCod,4,2) As cCodAge " _
'        & " FROM MovCol MC JOIN " & lsTablaDiaria & " M ON M.nMovNro = MC.nMovNro " _
'        & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser & lsFiltroAge _
'        & " AND SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And M.nMovFlag =" & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
'        & ") Mov ON Mov.cCodAge = Ag.cAgeCod   " & IIf(pnTipoGrupo = gGruposIngEgreAgeLocal, " And cAgeCod = '" & psCodAge & "'", " And cAgeCod <> '" & psCodAge & "' ") & ") " _
'        & " ORDER BY nAgeOrd "
'End If

    sSql = "sp_sel_GetOpeIngEgreColocaciones " & pnTipoGrupo & ", '" _
                                               & Format(pdFecha, "yyyymmdd") & "', '" _
                                               & psCodAge & "', '" _
                                               & psCodUser & "', " _
                                               & pnMoneda & ", '" _
                                               & Format(pdFechaDia, "yyyymmdd") & "', '" _
                                               & pgsCodUser & "', 0"

Set rs = oConect.CargaRecordSet(sSql)
Set GetOpeIngEgreColocaciones = rs
'If pdFecha <> pdFechaDia Then
'        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
'End If
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
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String) As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim lsFiltroUser2 As String
Dim lsFiltroAge2 As String
Dim oConect As COMConecta.DCOMConecta
Dim lsTablaDiaria As String

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function


'RIRO20150608, COMENTADO ****************************
'If pdFecha = pdFechaDia Then
'    lsTablaDiaria = " MovDiario "
'Else
'    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
'    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
'End If
'
'Set Rs = New adodb.Recordset
'If psCodAge <> "" Then
'    lsFiltroAge = " AND  SUBSTRING(M.CMOVNRO,18,2) = '" & psCodAge & "'"
'    lsFiltroAge2 = " AND  SUBSTRING(CMOVNRO,18,2) = '" & psCodAge & "'"
'End If
'If psCodUser <> "" Then
'    lsFiltroUser = " AND RIGHT(M.CMOVNRO,4) = '" & psCodUser & "' "
'    lsFiltroUser2 = " AND RIGHT(CMOVNRO,4) = '" & psCodUser & "' "
'End If
'
''--Otras Operaciones
'Sql = "SELECT cGrupoNombre, cCodAge = '" & psCodAge & "', SUM(nTotalMov) AS nTotalMov, " _
'    & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN EFECTIVO ELSE EFECTIVO*-1 END) AS EFECTIVO, " _
'    & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN CHEQUE ELSE CHEQUE*-1 END) AS CHEQUE, " _
'    & "  SUM(CASE WHEN G.cIngEgr = 'I' THEN ORDENPAGO ELSE ORDENPAGO*-1 END) AS ORDENPAGO, GO.cGrupoCod " _
'    & "FROM GRUPOOPE  AS G JOIN GRUPOSOPE GO ON GO.cGrupoCod=G.cGrupoCod " _
'    & "INNER JOIN (SELECT M.cOpeCod, O.COPEDESC, COUNT(*) AS nTotalMov, " _
'    & "  CASE WHEN MD.nDocTpo IS NULL THEN  SUM(MV.nMovImporte) ELSE 0 END AS EFECTIVO, " _
'    & "  CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN  SUM(MV.nMovImporte) ELSE 0 END AS CHEQUE, " _
'    & "  CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN  SUM(MV.nMovImporte) ELSE 0 END AS ORDENPAGO " _
'    & "  FROM " & lsTablaDiaria & " M JOIN MOVOPEVARIAS MV ON MV.nMovNro = M.nMovNro " _
'    & "  LEFT JOIN MOVDOC MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
'    & "  JOIN OPETPO O ON O.cOpeCod = M.cOpeCod " _
'    & "  WHERE   SUBSTRING(M.CMOVNRO,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser _
'    & "  AND M.NMOVFLAG =" & gMovFlagVigente & " AND MV.nMoneda =" & pnMoneda & " " _
'    & "  GROUP BY  M.cOpeCod, O.COPEDESC, MD.nDocTpo ) AS OPESERV " _
'    & "  ON OPESERV.cOpeCod = GO.cOpeCod " _
'    & " WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%'  " _
'    & " GROUP BY cGrupoNombre, GO.cGrupoCod "

'END RIRO *******************************************
    
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

sql = "stp_sel_OtrasOpe '" & Format(pdFecha, "yyyymmdd") & "', '" & psCodAge & "', '" & psCodUser & "', " & pnMoneda & ", '" & Format(pdFechaDia, "yyyymmdd") & "'"

Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreOtrasOpe = rs

'RIRO 20150608 ERS162-2014 *********************************
'If pdFecha <> pdFechaDia Then
'        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
'End If
'END RIRO **************************************************
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetOpeIngEgreServicios(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String) As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
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

Set rs = New adodb.Recordset
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
    & "  Cheque = 0, OrdenPago = 0 FROM " & lsTablaDiaria & " M JOIN MovServicios MV ON MV.nMovNro = M.nMovNro and MV.cOpecod <> '" & gITFGiroCancelEfect & "'" _
    & "  JOIN OPETPO O ON O.cOpeCod = M.cOpeCod " _
    & "  WHERE SUBSTRING(M.CMOVNRO,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser _
    & "  AND M.NMOVFLAG = " & gMovFlagVigente & " And MV.nMoneda = " & pnMoneda & " " _
    & "  GROUP BY  M.cOpeCod, O.COPEDESC) AS OPESERV " _
    & "  ON OPESERV.cOpeCod = GO.cOpeCod " _
    & " WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%'  " _
    & " GROUP BY cGrupoNombre, GO.cGrupoCod "
    'NAGL 20190307 Agregó and MV.cOpecod <> '" & gITFGiroCancelEfect & "'"
Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreServicios = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetOpeIngEgreSobranteFaltante(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String, _
                Optional ByVal psUsuarioBOVEDA As String = "BOVE") As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
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

Set rs = New adodb.Recordset
If psCodAge <> "" Then
    lsFiltroAge = " AND  SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "'"
End If

If psCodUser <> "" And psCodUser <> psUsuarioBOVEDA Then
    lsFiltroUser = " AND Right(M.cMovNro,4) = '" & psCodUser & "' And M.cOpeCod NOT IN ('" & COMDConstSistema.gOpeBoveAgeRegFaltante & "','" & COMDConstSistema.gOpeBoveAgeRegSobrante & "')"
ElseIf psCodUser = psUsuarioBOVEDA Then
    lsFiltroUser = " AND M.cOpeCod IN ('" & COMDConstSistema.gOpeBoveAgeRegFaltante & "','" & COMDConstSistema.gOpeBoveAgeRegSobrante & "') "
End If

'--Otras Operaciones
sql = "SELECT cGrupoNombre, SUM(nTotalMov) AS nTotalMov, cCodAge = '" & psCodAge & "', " _
    & "SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
    & "Cheque = 0, OrdenPago = 0, GO.cGrupoCod " _
    & " FROM GrupoOpe AS G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
    & " JOIN (SELECT  M.cOpeCod, O.cOpeDesc, COUNT(*) AS nTotalMov, SUM(Abs(MV.nMovImporte)) As Efectivo " _
    & " FROM  " & lsTablaDiaria & " M JOIN MovOpeVarias MV ON MV.nMovNro = M.nMovNro " _
    & " JOIN OPETPO O ON O.cOpeCod = M.cOpeCod " _
    & " WHERE SUBSTRING(M.cMovNro,1,8) ='" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser _
    & " AND M.nMovFlag = " & gMovFlagVigente & " AND MV.nMoneda = " & pnMoneda & " AND ABS(MV.nMovImporte) > 0 " _
    & " GROUP BY  M.cOpeCod, O.cOpeDesc ) AS OpeServ " _
    & " ON OpeServ.cOpeCod = GO.cOpeCod " _
    & " WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%'  " _
    & " GROUP BY cGrupoNombre, GO.cGrupoCod "

Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreSobranteFaltante = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function

'madm 20110127
Public Function GetOpeIngEgreSobranteFaltantePreCuadre(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String, _
                Optional ByVal psUsuarioBOVEDA As String = "BOVE", Optional ByVal pbind As Boolean = False) As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim oConect As COMConecta.DCOMConecta
Dim lsTablaDiaria As String
Dim lsFiltroPre As String

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

If pbind Then
     lsFiltroPre = " and nMovNro > (Select max(E.nMovNro) nMovNro From  MovDiario  M INNER JOIN MOv E ON M.nMovNro = E.nMovNro  "
     lsFiltroPre = lsFiltroPre & " Where SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' AND M.nMovFlag NOT IN (2,1) and SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "'" & lsFiltroUser & " AND M.cOpeCod in ('901018')  ) "
  End If

If pdFecha = pdFechaDia Then
    lsTablaDiaria = " MovDiario "
Else
    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
End If

Set rs = New adodb.Recordset
If psCodAge <> "" Then
    lsFiltroAge = " AND  SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "'"
End If

If psCodUser <> "" And psCodUser <> psUsuarioBOVEDA Then
    lsFiltroUser = " AND Right(M.cMovNro,4) = '" & psCodUser & "' And M.cOpeCod IN ('901041','901042')"
ElseIf psCodUser = psUsuarioBOVEDA Then
    lsFiltroUser = " AND M.cOpeCod IN ('" & COMDConstSistema.gOpeBoveAgeRegFaltante & "','" & COMDConstSistema.gOpeBoveAgeRegSobrante & "') "
End If

sql = "SELECT cGrupoNombre, SUM(nTotalMov) AS nTotalMov, cCodAge = '" & psCodAge & "', " _
    & "SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
    & "Cheque = 0, OrdenPago = 0, GO.cGrupoCod " _
    & " FROM GrupoOpe AS G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
    & " JOIN (SELECT  M.cOpeCod, O.cOpeDesc, COUNT(*) AS nTotalMov, SUM(Abs(MV.nMovImporte)) As Efectivo " _
    & " FROM  " & lsTablaDiaria & " M JOIN MovOpeVarias MV ON MV.nMovNro = M.nMovNro " _
    & " JOIN OPETPO O ON O.cOpeCod = M.cOpeCod " _
    & " WHERE SUBSTRING(M.cMovNro,1,8) ='" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser _
    & " AND M.nMovFlag = " & gMovFlagVigente & " AND MV.nMoneda = " & pnMoneda & " AND ABS(MV.nMovImporte) > 0 " _
    & " GROUP BY  M.cOpeCod, O.cOpeDesc ) AS OpeServ " _
    & " ON OpeServ.cOpeCod = GO.cOpeCod " _
    & " WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%'  " _
    & " GROUP BY cGrupoNombre, GO.cGrupoCod "

Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreSobranteFaltantePreCuadre = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetOpeIngEgreCompraVentaPreCuadre(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                        ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String, Optional ByVal pbind As Boolean = False) As adodb.Recordset

Dim sql As String
Dim rs As adodb.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim oConect As COMConecta.DCOMConecta
Dim lsTablaDiaria As String
Dim lsFiltroPre As String

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

If pbind Then
     lsFiltroPre = " and M.nMovNro > (Select max(E.nMovNro) nMovNro From  MovDiario  M INNER JOIN MOv E ON M.nMovNro = E.nMovNro  "
     lsFiltroPre = lsFiltroPre & " Where SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' AND M.nMovFlag NOT IN (2,1) and SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "'" & lsFiltroUser & " AND M.cOpeCod in ('901018')  ) "
  End If
  
If pdFecha = pdFechaDia Then
    lsTablaDiaria = " MovDiario "
Else
    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
End If

Set rs = New adodb.Recordset
If psCodAge <> "" Then
    lsFiltroAge = " AND  SUBSTRING(M.CMOVNRO,18,2) ='" & psCodAge & "'"
End If
If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.CMOVNRO,4) = '" & psCodUser & "' "
End If

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
    & "        WHERE SUBSTRING(M.CMOVNRO,1,8) ='" & Format$(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser & lsFiltroPre _
    & "        AND M.NMOVFLAG = " & gMovFlagVigente & "" _
    & "     GROUP BY M.cOpeCod, O.cOpeDesc, TC.nMovTpoCambio) AS MTC ) AS TC) AS MTC " _
    & " ON MTC.cOpeCod = GO.cOpeCod Group by G.cGrupoNombre, GO.cGrupoCod"

Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreCompraVentaPreCuadre = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function
'end madm

Public Function GetOpeIngEgreCompraVenta(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                        ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String) As adodb.Recordset

Dim sql As String
Dim rs As adodb.Recordset
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

Set rs = New adodb.Recordset
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
                ByVal psCodAge As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String) As adodb.Recordset

Dim sSql As String
Dim rs As adodb.Recordset
'Dim lsFiltroAge As String
Dim oConect As COMConecta.DCOMConecta
'Dim lsTablaDiaria As String

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

'If pdFecha = pdFechaDia Then
'    lsTablaDiaria = " MovDiario "
'Else
'    Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
'    lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
'End If

Set rs = New adodb.Recordset


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

'sSql = "SELECT cGrupoNombre, GO.cGrupoCod, SUM(nTotalMov) nTotalMov, cCodAge = '" & psCodAge & "', " _
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
'sSql = sSql & " UNION SELECT cOpeCod, " _
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

sSql = "EXEC stp_sel_GetOpeIngEgreHabDev '" & Format(pdFechaDia, "yyyymmdd") & "'," & pnTipoGrupo & ",'" & Format(pdFecha, "yyyymmdd") & "','" & psCodAge & "'," & pnMoneda & ",'" & pgsCodUser & "'"  'EJVG20160309 Se corrigió para mostrar las confirmaciones del día que fueron habilitados en días anteriores

Set rs = oConect.CargaRecordSet(sSql)
Set GetOpeIngEgreHabDev = rs
'If pdFecha <> pdFechaDia Then
'        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
'End If
oConect.CierraConexion: Set oConect = Nothing
End Function

'madm 20110115
Public Function GetOpePreCuadreOperaciones(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String, _
                Optional ByVal psUsuarioBOVEDA As String = "BOVE") As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
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

Set rs = New adodb.Recordset
If psCodAge <> "" Then
    lsFiltroAge = " AND  SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "'"
End If

If psCodUser <> "" And psCodUser <> psUsuarioBOVEDA Then
    lsFiltroUser = " AND Right(M.cMovNro,4) = '" & psCodUser & "' And M.cOpeCod NOT IN ('" & COMDConstSistema.gOpeBoveAgeRegFaltante & "','" & COMDConstSistema.gOpeBoveAgeRegSobrante & "')"
ElseIf psCodUser = psUsuarioBOVEDA Then
    lsFiltroUser = " AND M.cOpeCod IN ('" & COMDConstSistema.gOpeBoveAgeRegFaltante & "','" & COMDConstSistema.gOpeBoveAgeRegSobrante & "') "
End If

If pnMoneda = gMonedaNacional Then
sql = "SELECT cGrupoNombre, SUM(nTotalMov) AS nTotalMov, cCodAge = '" & psCodAge & "', " _
    & "SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
    & "Cheque = 0, OrdenPago = 0, GO.cGrupoCod " _
    & " FROM GrupoOpe AS G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
    & " JOIN (SELECT  M.cOpeCod, O.cOpeDesc, COUNT(*) AS nTotalMov, SUM(Abs(MV.nMontoSol)) As Efectivo " _
    & " FROM  " & lsTablaDiaria & " M JOIN MovPreCuadre MV ON MV.nMovNro = M.nMovNro " _
    & " JOIN OPETPO O ON O.cOpeCod = M.cOpeCod " _
    & " WHERE SUBSTRING(M.cMovNro,1,8) ='" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser _
    & " AND M.nMovFlag = " & gMovFlagVigente & " AND ABS(MV.nMontoSol) > 0 " _
    & " GROUP BY  M.cOpeCod, O.cOpeDesc ) AS OpeServ " _
    & " ON OpeServ.cOpeCod = GO.cOpeCod " _
    & " WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%'  " _
    & " GROUP BY cGrupoNombre, GO.cGrupoCod "
Else
sql = "SELECT cGrupoNombre, SUM(nTotalMov) AS nTotalMov, cCodAge = '" & psCodAge & "', " _
    & "SUM(CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
    & "Cheque = 0, OrdenPago = 0, GO.cGrupoCod " _
    & " FROM GrupoOpe AS G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
    & " JOIN (SELECT  M.cOpeCod, O.cOpeDesc, COUNT(*) AS nTotalMov, SUM(Abs(MV.nMontoDol)) As Efectivo " _
    & " FROM  " & lsTablaDiaria & " M JOIN MovPreCuadre MV ON MV.nMovNro = M.nMovNro " _
    & " JOIN OPETPO O ON O.cOpeCod = M.cOpeCod " _
    & " WHERE SUBSTRING(M.cMovNro,1,8) ='" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser _
    & " AND M.nMovFlag = " & gMovFlagVigente & " AND ABS(MV.nMontoDol) > 0 " _
    & " GROUP BY  M.cOpeCod, O.cOpeDesc ) AS OpeServ " _
    & " ON OpeServ.cOpeCod = GO.cOpeCod " _
    & " WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%'  " _
    & " GROUP BY cGrupoNombre, GO.cGrupoCod "
End If

Set rs = oConect.CargaRecordSet(sql)
Set GetOpePreCuadreOperaciones = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetOpeIngEgreHabDevCajeroPreCuadre(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String) As adodb.Recordset

Dim sSql As String
Dim rs As adodb.Recordset
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

Set rs = New adodb.Recordset


sSql = "SELECT cGrupoNombre, GO.cGrupoCod, Sum(nTotalMov) nTotalMov, cCodAge = '" & psCodAge & "', " _
    & "Efectivo = CASE WHEN G.cIngEgr = 'I' THEN SUM(nMontoHabDev) ELSE SUM(nMontoHabDev)*-1 END, " _
    & "SUM(Cheque) As Cheque, SUM(OrdenPago) As OrdenPago FROM GrupoOpe G JOIN GruposOpe GO " _
    & "ON GO.cGrupoCod = G.cGrupoCod JOIN " _
    & "(SELECT cOpeCod, nTotalMov, Efectivo AS nMontoHabDev, 0 AS Cheque , 0 AS OrdenPago FROM " _
    & "(SELECT M1.cOpeCod, COUNT(*) AS nTotalMov, SUM(MH.nMovImporte) AS Efectivo " _
    & "FROM OpeTpo O JOIN " & lsTablaDiaria & " M JOIN MovHabDevCajeroPreCuadre MH ON MH.nMovNro = M.nMovNro ON O.cOpeCod = M.cOpeCod " _
    & "JOIN MovRef MR JOIN  " & lsTablaDiaria & "  M1 ON MR.nMovNro = M1.nMovNro ON MR.nMovNroRef = M.nMovNro " _
    & "WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format$(pdFecha, "yyyymmdd") & "' AND MH.cUsuDest = '" & psCodUser & "' AND " _
    & "M.nMovFlag = 0 AND MH.cAgeCod = '" & psCodAge & "' And MH.nMoneda = " & pnMoneda & " And M1.nMovFlag = " & gMovFlagVigente & " " _
    & "Group by M1.cOpeCod UNION " _
    & "SELECT M.cOpeCod, COUNT(*) AS nTotalMov, SUM(MH.nMovImporte) AS Efectivo " _
    & "FROM OpeTpo O JOIN " & lsTablaDiaria & " M JOIN MovHabDevCajeroPreCuadre MH ON MH.nMovNro = M.nMovNro ON O.cOpeCod = M.cOpeCod " _
    & "JOIN MovRef MR JOIN " & lsTablaDiaria & " M1 ON MR.nMovNro = M1.nMovNro ON MR.nMovNroRef = M.nMovNro " _
    & "WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format$(pdFecha, "yyyymmdd") & "' AND MH.cUsuOrig = '" & psCodUser & "' AND " _
    & "M.nMovFlag = 0 AND MH.cAgeCod = '" & psCodAge & "' And MH.nMoneda = " & pnMoneda & " And M1.nMovFlag = " & gMovFlagVigente & " " _
    & "Group by M.cOpeCod) HabDev) OpeHD ON OpeHD.cOpeCod = GO.cOpeCod Where " _
    & "G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%' GROUP BY cGrupoNombre, GO.cGrupoCod, G.cIngEgr"

Set rs = oConect.CargaRecordSet(sSql)
Set GetOpeIngEgreHabDevCajeroPreCuadre = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If

oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetOpeIngEgreHabDevBillCajeroPreCuadre(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String) As adodb.Recordset

Dim sSql As String
Dim rs As adodb.Recordset
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

Set rs = New adodb.Recordset


sSql = "SELECT cGrupoNombre, GO.cGrupoCod, Sum(nTotalMov) nTotalMov, cCodAge = '" & psCodAge & "', " _
    & "Efectivo = CASE WHEN G.cIngEgr = 'I' THEN SUM(nMontoHabDev) ELSE SUM(nMontoHabDev)*-1 END, " _
    & "SUM(Cheque) As Cheque, SUM(OrdenPago) As OrdenPago FROM GrupoOpe G JOIN GruposOpe GO " _
    & "ON GO.cGrupoCod = G.cGrupoCod JOIN " _
    & "(SELECT cOpeCod, nTotalMov, Efectivo AS nMontoHabDev, 0 AS Cheque , 0 AS OrdenPago FROM " _
    & "(SELECT M.cOpeCod, COUNT(*) AS nTotalMov, SUM(MH.nMovImporte) AS Efectivo " _
    & "FROM OpeTpo O JOIN " & lsTablaDiaria & " M JOIN MovHabDevCajeroPreCuadre MH ON MH.nMovNro = M.nMovNro ON O.cOpeCod = M.cOpeCod " _
    & "WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format$(pdFecha, "yyyymmdd") & "' AND MH.cUsuOrig = '" & psCodUser & "' AND " _
    & "M.nMovFlag = 0 AND MH.cAgeCod = '" & psCodAge & "' And MH.nMoneda = " & pnMoneda & " " _
    & "Group by M.cOpeCod) HabDev) OpeHD ON OpeHD.cOpeCod = GO.cOpeCod Where " _
    & "G.cGrupoCod IN ('" & oVar.gsCajDevBilletaje & "') GROUP BY cGrupoNombre, GO.cGrupoCod, G.cIngEgr"

Set rs = oConect.CargaRecordSet(sSql)
Set GetOpeIngEgreHabDevBillCajeroPreCuadre = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function
'end madm

Public Function GetOpeIngEgreHabDevCajero(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String) As adodb.Recordset

Dim sSql As String
Dim rs As adodb.Recordset
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

Set rs = New adodb.Recordset


sSql = "SELECT cGrupoNombre, GO.cGrupoCod, Sum(nTotalMov) nTotalMov, cCodAge = '" & psCodAge & "', " _
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

Set rs = oConect.CargaRecordSet(sSql)
Set GetOpeIngEgreHabDevCajero = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If

oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetOpeIngEgreHabDevBillCajero(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String) As adodb.Recordset

Dim sSql As String
Dim rs As adodb.Recordset
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

Set rs = New adodb.Recordset


sSql = "SELECT cGrupoNombre, GO.cGrupoCod, Sum(nTotalMov) nTotalMov, cCodAge = '" & psCodAge & "', " _
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

Set rs = oConect.CargaRecordSet(sSql)
Set GetOpeIngEgreHabDevBillCajero = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetCompraVentaRepo(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                              ByVal psCodUser As String, ByVal psCodAge As String, Optional lbHistorica As Boolean = False) As adodb.Recordset
Dim sql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Set rs = New adodb.Recordset
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

Public Function GetFaltanteCajero(ByVal psOpeCod As String, ByVal psCodUser As String, ByVal psCodAge As String) As adodb.Recordset
Dim sql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset
Dim lsOpeCod As String

Set rs = New adodb.Recordset
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
Dim oMov As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New COMDMov.DCOMMov

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

Public Function GrabaRegistroSobranteFaltante(ByVal sMovNro As String, ByVal sOpeCod As String, _
        ByVal sMovDesc As String, ByVal nImporte As Double, ByVal nMoneda As Moneda) As Integer
Dim oMov As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim nMovNro As Long
Set oMov = New COMDMov.DCOMMov

On Error GoTo GrabaRegularizaFaltanteErr
oMov.BeginTrans
lbTrans = True
GrabaRegistroSobranteFaltante = 1
oMov.InsertaMov sMovNro, sOpeCod, sMovDesc
nMovNro = oMov.GetnMovNro(sMovNro)
oMov.InsertaMovOpeVarias nMovNro, "", Trim(sMovDesc), nImporte, nMoneda
oMov.CommitTrans
lbTrans = False
GrabaRegistroSobranteFaltante = 0
Set oMov = Nothing
Exit Function
GrabaRegularizaFaltanteErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ImprimeBilletaje(ByVal rsBillMN As adodb.Recordset, ByVal rsBillME As adodb.Recordset, _
        ByVal rsMonMN As adodb.Recordset, ByVal rsMonME As adodb.Recordset, _
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

sCad = sCad & oImpre.CabeRepoCaptac("", "", nCarLin, "SECCION OPERACIONES", sTitRp1, sTitRp2, sMoneda, "1", Trim(sNomAge), dFecSis, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea

'''sMoneda = "SOLES" 'marg ers044-2016
sMoneda = StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) 'marg ers044-2016
sMonedaTit = "MONEDA NACIONAL"
sSimbolo = "S/."
sCad = sCad & sMonedaTit & oImp.gPrnSaltoLinea
sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
sCad = sCad & ImprimeRsBilletaje(rsBillMN, nTotalBill)
RSet sTotalMonto = Format$(nTotalBill, "#,##0.00")
sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
sCad = sCad & "SUB TOTAL BILLETAJE" & Space(23) & sTotalMonto & oImp.gPrnSaltoLinea
sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
sCad = sCad & ImprimeRsBilletaje(rsMonMN, nTotalMon)
RSet sTotalMonto = Format$(nTotalMon, "#,##0.00")
sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
sCad = sCad & "SUB TOTAL MONEDA" & Space(26) & sTotalMonto & oImp.gPrnSaltoLinea
sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
RSet sTotalMonto = sSimbolo & " " & Format$(nTotalBill + nTotalMon, "#,##0.00")
sCad = sCad & "TOTAL " & sMonedaTit & Space(21) & sTotalMonto & oImp.gPrnSaltoLinea
sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
sCad = sCad & oImp.gPrnSaltoLinea

sMoneda = "DOLARES"
sMonedaTit = "MONEDA EXTRANJERA"
sSimbolo = "US$"
sCad = sCad & sMonedaTit & oImp.gPrnSaltoLinea
sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
sCad = sCad & ImprimeRsBilletaje(rsBillME, nTotalBill)
RSet sTotalMonto = Format$(nTotalBill, "#,##0.00")
sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
sCad = sCad & "SUB TOTAL BILLETAJE" & Space(23) & sTotalMonto & oImp.gPrnSaltoLinea
sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
'sCad = sCad & ImprimeRsBilletaje(rsMonME, nTotalMon)
'RSet sTotalMonto = Format$(nTotalMon, "#,##0.00")
'sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
'sCad = sCad & "SUB TOTAL MONEDA" & Space(26) & sTotalMonto & oImp.gPrnSaltoLinea
'sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
'RSet sTotalMonto = sSimbolo & " " & Format$(nTotalBill + nTotalMon, "#,##0.00")
sCad = sCad & "TOTAL " & sMonedaTit & Space(21) & sTotalMonto & oImp.gPrnSaltoLinea
sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
ImprimeBilletaje = sCad
End Function

Private Function ImprimeRsBilletaje(ByVal rs As adodb.Recordset, ByRef nTotalMonto As Double) As String
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
        sCad = sCad & sDescripcion & Space(2) & sCantidad & Space(2) & sMonto & oImp.gPrnSaltoLinea
        rs.MoveNext
    Loop
End If
ImprimeRsBilletaje = sCad
End Function

Public Function YaRealizoDevBilletaje(ByVal sUsuario As String, ByVal dFecha As String, Optional ByVal psCodAge As String = "") As Boolean
Dim sSql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset

Set rs = New adodb.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

sSql = "Select M.cOpeCod From Mov M INNER JOIN MovHabDevCajero H ON M.nMovNro = H.nMovNro " _
    & "Where M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And cOpeCod = '" & gOpeHabCajDevBilletaje & "'" _
    & "And H.cUsuOrig = '" & sUsuario & "' And M.nMovFlag IN (" & gMovFlagVigente & ")"

If psCodAge <> "" Then
    sSql = sSql & " AND Substring(cMovNro, 18,2) = '" & psCodAge & "'"
End If

Set rs = oConect.CargaRecordSet(sSql)
If rs.EOF And rs.BOF Then
    YaRealizoDevBilletaje = False
Else
    YaRealizoDevBilletaje = True
End If
rs.Close
Set rs = Nothing
Set oConect = Nothing
End Function

'MIOL 20120727, CAMBIO BOBEDA REGISTRO EFECTIVO ******
Public Function YaRealizoRegEfectivo(ByVal sUsuario As String, ByVal dFecha As String, Optional ByVal psCodAge As String = "") As Boolean
Dim sSql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset

Set rs = New adodb.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

sSql = "Select M.cOpeCod From Mov M " _
    & "Where M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%" & sUsuario & "' And cOpeCod = '" & gOpeBoveAgeRegEfect & "'" _
    & "And M.nMovFlag IN (" & gMovFlagVigente & ")"

If psCodAge <> "" Then
    sSql = sSql & " AND Substring(cMovNro, 18,2) = '" & psCodAge & "'"
End If

Set rs = oConect.CargaRecordSet(sSql)
If rs.EOF And rs.BOF Then
    YaRealizoRegEfectivo = False
Else
    YaRealizoRegEfectivo = True
End If
rs.Close
Set rs = Nothing
Set oConect = Nothing
End Function
'END MIOL ********************************************

Public Function GetDetalleHabDevCajero(ByVal sAreaCod As String, ByVal sAgecod As String, _
        ByVal dFecha As Date, Optional ByVal sCajero As String = "", Optional nMoneda As Moneda = gMonedaNacional) As adodb.Recordset


Dim sql As String
Dim rs As adodb.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sFiltro As String

Set oCon = New COMConecta.DCOMConecta
Set rs = New adodb.Recordset
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
Dim sSql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset

sSql = "Select M.nMovNro From MovUserEfectivo E JOIN Mov M ON E.nMovNro = M.nMovNro Where " _
    & "M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And E.cUser = '" & sUsuario & "' " _
    & "And SUBSTRING(M.cMovNro,18,2) = '" & sAgencia & "' And M.cOpeCod IN ('" & sOperacion & "')"

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
Set rs = oConect.CargaRecordSet(sSql)
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
                              ByVal psCodUser As String, ByVal psCodAge As String) As adodb.Recordset
Dim sSql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset
Dim lsFiltroUser As String

Set rs = New adodb.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
If psCodUser <> "" Then
    lsFiltroUser = "  AND RIGHT(M.cMovNro,4)='" & psCodUser & "' "
End If

sSql = "SELECT CONVERT(Varchar(10),CONVERT(DATETIME, SUBSTRING(M.cMovNro,1,8)),103) + ' ' + " _
    & "SUBSTRING(M.cMovNro,9,2) + ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) Fecha, " _
    & "O.cOpeDesc, RIGHT(M.cMovNro,4) Usu, F.cctacod, cSimbolo = 'S/.', MS.nMonto, M.nMovNro, " _
    & "nMoneda = 1, M.cMovDesc FROM OpeTpo O JOIN Mov M JOIN MovServicios MS ON MS.nMovNro = M.nMovNro " _
    & "ON O.cOpeCod = M.cOpeCod INNER JOIN FONCODESMOV FM ON FM.nNroMov = MS.nMovNro " _
    & " INNER JOIN FONCODES F ON FM.cCtaCod = F.cCtaCod WHERE M.nMovFlag = " & gMovFlagVigente & " AND " _
    & "M.cOpeCod = '" & psOpeCod & "' And SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " & lsFiltroUser _
    & "AND SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' And '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "ORDER BY M.nMovNro "
 
Set rs = oConect.CargaRecordSet(sSql)
Set GetMovFoncodes = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
'PLAN BICI
'CRSF - 24/06
Public Function GetMovPlanBici(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                              ByVal psCodUser As String, ByVal psCodAge As String) As adodb.Recordset
Dim sSql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset
Dim lsFiltroUser As String

Set rs = New adodb.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
If psCodUser <> "" Then
    lsFiltroUser = "  AND RIGHT(M.cMovNro,4)='" & psCodUser & "' "
End If

sSql = "SELECT CONVERT(Varchar(10),CONVERT(DATETIME, SUBSTRING(M.cMovNro,1,8)),103) + ' ' + " _
    & "SUBSTRING(M.cMovNro,9,2) + ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2) Fecha, " _
    & "O.cOpeDesc, RIGHT(M.cMovNro,4) Usu, F.cctacod, cSimbolo = 'S/.', MS.nMonto, M.nMovNro, " _
    & "nMoneda = 1, M.cMovDesc FROM OpeTpo O JOIN Mov M JOIN MovServicios MS ON MS.nMovNro = M.nMovNro " _
    & "ON O.cOpeCod = M.cOpeCod INNER JOIN PLANBICIMOV FM ON FM.nNroMov = MS.nMovNro " _
    & " INNER JOIN PLANBICI F ON FM.cCtaCod = F.cCtaCod WHERE M.nMovFlag = " & gMovFlagVigente & " AND " _
    & "M.cOpeCod = '" & psOpeCod & "' And SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " & lsFiltroUser _
    & "AND SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' And '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "ORDER BY M.nMovNro "
 
Set rs = oConect.CargaRecordSet(sSql)
Set GetMovPlanBici = rs
oConect.CierraConexion: Set oConect = Nothing
End Function


Public Function GetOpeIngEgreOpeCMACs(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, ByVal psCodAge As String, _
            ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String) As adodb.Recordset

Dim sSql As String
Dim rs As adodb.Recordset
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
    
    
Set rs = New adodb.Recordset
If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.cMovNro,4)  = '" & psCodUser & "' "
End If



sSql = "SELECT (OpeCol.cCodAge + '2') As nAgeOrd, OpeCol.cCodAge, cGrupoNombre, SUM(nTotalMov) AS nTotalMov, " _
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
sSql = sSql + " UNION " _
    & "(SELECT DISTINCT (P.cPersCod + '1') As nAgeOrd, P.cPersCod As cAgeCod, P.cPersNombre As cAgeDescripcion,0,0,0,0,'' " _
    & "FROM Persona P JOIN " _
    & "(SELECT MC.cPersCod As cCodAge " _
    & " FROM MovCMAC MC JOIN " & lsTablaDiaria & " M ON M.nMovNro = MC.nMovNro " _
    & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
    & " AND MC.nMoneda = '" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
    & ") Mov ON Mov.cCodAge = P.cPersCod) " _
    & " ORDER BY nAgeOrd "

Set rs = oConect.CargaRecordSet(sSql)
Set GetOpeIngEgreOpeCMACs = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GrabaCierreAgencia(ByVal sMovNro As String, ByVal sOpeCod As String, ByVal sMovDesc As String) As Integer
Dim oMov As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New COMDMov.DCOMMov

On Error GoTo Error
oMov.BeginTrans
lbTrans = True
GrabaCierreAgencia = 1
oMov.InsertaMov sMovNro, sOpeCod, sMovDesc, gMovEstContabNoContable, gMovFlagVigente
oMov.CommitTrans
lbTrans = False
GrabaCierreAgencia = 0
Set oMov = Nothing
Exit Function
Error:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function YaRealizoCierreAgencia(ByVal sCodAge As String, ByVal dFecha As Date) As Boolean
    Dim sSql As String
    Dim oConect As COMConecta.DCOMConecta
    Dim rs As adodb.Recordset

    Set rs = New adodb.Recordset
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    '***Modificado por ELRO el 20120312 según Acta N° 045-2012/TI-D
    'sSql = "Select M.cOpeCod From MovDiario M " _
    '    & "Where M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And M.cOpeCod = '" & gOpeCajaCierreAgencia & "'" _
    '    & "And M.nMovFlag IN (" & gMovFlagVigente & ") AND Substring(M.cMovNro, 18,2) = '" & sCodage & "'"
    sSql = "stp_sel_ValidarMovDiarioAgencia '" & sCodAge & "', '" & Format$(dFecha, "yyyymmdd") & "', '" & gOpeCajaCierreAgencia & "'"
    '***Modificado por ELRO****************************************

    Set rs = oConect.CargaRecordSet(sSql)
    If rs.EOF And rs.BOF Then
        YaRealizoCierreAgencia = False
    Else
        YaRealizoCierreAgencia = True
    End If
    rs.Close
    Set rs = Nothing
    Set oConect = Nothing
End Function

Public Function YaRegistroEfectivoAutomatico(ByVal sCodAge As String, ByVal dFecha As Date) As Integer
Dim sSql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As adodb.Recordset

Set rs = New adodb.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

sSql = "select cCodAge From EfectivoAutomatico "
sSql = sSql & " where convert(varchar(8), dfecha, 112)='" & Format(dFecha, "YYYYMMdd") & "' "
sSql = sSql & " And cCodAge='" & Trim(sCodAge) & "' and nFlag IN (" & gMovFlagVigente & ")"

Set rs = oConect.CargaRecordSet(sSql)
If rs.EOF And rs.BOF Then
    
    rs.Close
    
    sSql = "Select M.cOpeCod From Mov M " _
    & "Where M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And M.cOpeCod = '" & gOpeCajaCierreAgencia & "'" _
    & "And M.nMovFlag IN (" & gMovFlagVigente & ") AND Substring(M.cMovNro, 18,2) = '" & sCodAge & "'"

    Set rs = oConect.CargaRecordSet(sSql)
    If rs.EOF And rs.BOF Then
        
        rs.Close
    
        sSql = "Select M.* from mov M Inner Join MovUserEfectivo MU On M.nMovNro=MU.nMovNro "
        sSql = sSql & " Where M.nMovFlag IN (" & gMovFlagVigente & ") AND Substring(M.cMovNro, 18,2) = '" & sCodAge & "' "
        sSql = sSql & " And M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And M.cOpeCod IN('" & gOpeBoveAgeRegEfect & "', '" & gOpeHabCajRegEfect & "') "
        
        Set rs = oConect.CargaRecordSet(sSql)
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

Public Sub RegistraEfectivoAutomatico(ByVal sCodAge As String, ByVal dFecha As Date, ByVal sCodUsu As String, ByVal psNomAge As String)
Dim oConect As COMConecta.DCOMConecta
Dim oMov As New COMDMov.DCOMMov
Dim ClsMov As COMNContabilidad.NCOMContFunciones
Dim reg As adodb.Recordset
Dim sSql As String

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
Set ClsMov = New COMNContabilidad.NCOMContFunciones

Set oMov = New COMDMov.DCOMMov

If oConect.AbreConexion = False Then Exit Sub

'sSql = "select distinct mu.nmovnro, m.cOpeCod,mu.cUser, mu.cEfectivoCod, mu.nMonto "
'sSql = sSql & " from Mov M "
'sSql = sSql & " Inner Join MovUserEfectivo MU "
'sSql = sSql & " on M.nMovNro=MU.nMovNro "
'sSql = sSql & " Where cOpeCod in('" & gOpeBoveAgeRegEfect & "', '" & gOpeHabCajRegEfect & "') "
'sSql = sSql & " and m.cmovnro like '" & Format(dfechaAnt, "YYYYMMdd") & "%' "
'sSql = sSql & " and m.nmovflag IN(" & gMovFlagVigente & ") AND Substring(M.cMovNro, 18,2) = '" & sCodAge & "' "
'sSql = sSql & " Order by mu.nmovnro, m.copecod "

sSql = "select distinct mu.cEfectivoCod, sum(mu.nMonto) as nMonto "
sSql = sSql & " from Mov M "
sSql = sSql & " Inner Join MovUserEfectivo MU "
sSql = sSql & " on M.nMovNro=MU.nMovNro "
sSql = sSql & " Where cOpeCod in('" & gOpeBoveAgeRegEfect & "', '" & gOpeHabCajRegEfect & "') "
sSql = sSql & " and m.cmovnro like '" & Format(dfechaAnt, "YYYYMMdd") & "%' "
sSql = sSql & " and m.nmovflag IN(" & gMovFlagVigente & ") AND Substring(M.cMovNro, 18,2) = '" & sCodAge & "' "
sSql = sSql & " GROUP BY mu.cEfectivoCod "

Set reg = oConect.CargaRecordSet(sSql)
oConect.CierraConexion
Set oConect = Nothing
If Not reg.EOF And Not reg.BOF Then
    lnTempo = 0
    oMov.BeginTrans
    
    sMovNro = ClsMov.GeneraMovNro(dFecha, sCodAge, "SIST")
    
    oMov.InsertaMov sMovNro, gOpeBoveAgeRegEfect, "REGISTRO DE EFECTIVO AUTOMATICO - " & sCodUsu, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(sMovNro)
    lsCodUser = "BOVE"
    
    Do While Not reg.EOF
        oMov.InsertaMovUserEfectivo lnMovNro, lsCodUser, reg!cEfectivoCod, reg!nMonto
        Sleep 120
        'If reg.Bookmark = 18 Then Stop
        reg.MoveNext
    Loop
    oMov.InsertaEfectivoAutomatico dFecha, sCodAge, sCodUsu
    '****************** Cierre de agencia EJRS ******************************
    sMovNro = ClsMov.GeneraMovNro(dFecha, sCodAge, "SIST")
    oMov.InsertaMov sMovNro, gOpeCajaCierreAgencia, "Cierre Caja Agencia Automatico : " & sCodAge & " - " & psNomAge, gMovEstContabNoContable, gMovFlagVigente
    '************************************************************************
    oMov.CommitTrans
End If
reg.Close
Set reg = Nothing

lbTrans = False

Set ClsMov = Nothing
Set oMov = Nothing

Exit Sub
ErrorRegistraEfectivoAutomatico:
    If lbTrans Then
        lbTrans = False
        oMov.RollbackTrans
    End If
    Err.Raise vbObjectError + 100, "RegistraEfectivoAutomatico", "Efectivo Automatico "

End Sub

Public Sub RegistraExtornoAutomatico(ByVal sCodAge As String, ByVal dFecha As Date, ByVal sCodUsu As String)
Dim oConect As COMConecta.DCOMConecta
Dim reg As adodb.Recordset
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
sSQL1 = sSQL1 & " and nmovflag IN(" & gMovFlagVigente & ") AND Substring(cMovNro, 18,2) = '" & sCodAge & "' "

sSql2 = " Update EfectivoAutomatico set nFlag=" & gMovFlagExtornado & " "
sSql2 = sSql2 & " where convert(varchar(8), dfecha, 112)='" & Format(dFecha, "YYYYMMdd") & "' and cCodAge='" & Trim(sCodAge) & "' "
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


Public Function GetUsuariosMovEfectivo(ByVal sAgencia As String, ByVal dFecha As Date) As adodb.Recordset
Dim sSql As String
Dim oConect As COMConecta.DCOMConecta
'ANDE 20171018
'sSql = "Select RIGHT(M.cMovNro,4) Usuario, P.cPersNombre From Persona P JOIN RRHH RH JOIN Mov M JOIN " _
'    & "(SELECT GO.cOpeCod From GrupoOpe G JOIN GruposOpe GO ON G.cGrupoCod = GO.cGrupoCod " _
'    & "WHERE G.nEfectivo = 1 And cOpeCod NOT LIKE '9010%' And cOpeCod NOT IN (401401, 402401, 401571, 402571) " _
'    & "And cOpeCod NOT IN (Select cOpeCod from OpeDoc Where nDocTpo = 47) ) O ON " _
'    & "M.cOpeCod = O.cOpeCod ON RH.cUser = RIGHT(M.cMovNro,4) ON P.cPersCod = RH.cPersCod " _
'    & "Where M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' " _
'    & "And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & sAgencia & "' " _
'    & "GROUP BY RIGHT(M.cMovNro,4), P.cPersNombre"

sSql = "exec stp_sel_GetUsuariosMovEfectivo '" & sAgencia & "','" & gMovFlagVigente & "','" & Format$(dFecha, "yyyymmdd") & "'"
'END ANDE
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
Set GetUsuariosMovEfectivo = oConect.CargaRecordSet(sSql)
Set oConect = Nothing
End Function

Public Function GetCajeroOpeIngxRef(ByVal psPersCod As String, ByVal pnMoneda As Moneda) As adodb.Recordset
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

Public Function GetSobFalMayUNO(ByVal psFecha As String, ByVal psCodAge As String) As adodb.Recordset
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

Public Function ImpCajeroIngEgre(ByVal rsMN As adodb.Recordset, ByVal rsME As adodb.Recordset, ByVal pdFecha As String, ByVal gsNomAge As String, ByVal psNroMovMN As String, ByVal psNroMovME As String, _
             ByVal psTotalIngresosN As String, ByVal psTotalEgresosN As String, ByVal psTotalHabilitacionN As String, ByVal psTotalDevolucionN As String, _
             ByVal psTotalDevBilletajeN As String, ByVal psTotalSobFaltN As String, ByVal psTotalPlanillaN As String, ByVal psSaldoAnteriorN As String, _
             ByVal psNetoMovimientosN As String, ByVal psSaldoFinalN As String, ByVal psDiferenciaN As String, _
             ByVal psTotalIngresosE As String, ByVal psTotalEgresosE As String, ByVal psTotalHabilitacionE As String, ByVal psTotalDevolucionE As String, _
             ByVal psTotalDevBilletajeE As String, ByVal psTotalSobFaltE As String, ByVal psTotalPlanillaE As String, ByVal psSaldoAnteriorE As String, _
             ByVal psNetoMovimientosE As String, ByVal psSaldoFinalE As String, ByVal psDiferenciaE As String, ByVal psUsuario As String, ByVal gdFecSis As String, _
             ByVal psNomCmac As String, ByVal psDescAge As String, Optional ByVal pbArqueo As Boolean = False, Optional ByVal psUserArqueo As String = "", _
             Optional ByVal psNumOpeMN As String = "", Optional ByVal psNumOpeME As String = "", Optional rsEgresos As adodb.Recordset) As String 'Add by EJVG 20110809 pnNumOpeMN y pnNumOpeME
    
    Dim nCarLin As Integer
    Dim nTotalEfectivo As Double, nTotalCheque As Double, nTotalOrdPag As Double, nTotalTotal As Double
    Dim nTotalNoEfectivo As Double 'FRHU ERS048-2014
    Dim sTitRp1 As String, sTitRp2 As String
    Dim sNumPag As String, sFchtra As String
    Dim nLinPag As Integer, nCntPag As Integer
    Dim oImpre As New COMFunciones.FCOMImpresion
    Dim oImpreCap As New COMNCaptaGenerales.NCOMCaptaImpresion
    Dim sNombre As String * 35, sNombre2 As String * 30
    Dim sEfectivo As String * 12, sCheques As String * 12, sOrdPag As String * 12, sTotFila As String * 12
    Dim sNoEfectivo As String * 12 'FRHU ERS048-2014
    Dim sNumReg As String * 5
    Dim sNetoMovimientos As String * 12, sNroMovimientos As String * 10, sSaldoFinal As String * 12
    Dim sMoneda As String
    Dim lsTextArqueo As String
    Dim sCad As String
    Dim i As Long
    
    sNumPag = ""
    nLinPag = 0
    'nCarLin = 98
    nCarLin = 112 'FRHU 20150602 ERS048-2014
    '*** PEAC 20090928
    If pbArqueo Then
        sTitRp1 = "ARQUEO REALIZADO AL USUARIO : " & psUserArqueo
    Else
        sTitRp1 = "RESUMEN DE INGRESOS Y EGRESOS - " & psUsuario
    End If
    '*** FIN PEAC
    sTitRp2 = "DEL " & Trim(pdFecha)
    nCntPag = 1
    sFchtra = Format$(gdFecSis, "dd/mm/yyyy")
    
    'Moneda Nacional
    '''sMoneda = "SOLES" 'marg ers044-2016
    sMoneda = StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) 'marg ers044-2016
    sCad = sCad & oImp.gPrnNegritaON
    sCad = sCad & oImpreCap.CabeRepoCaptac(psNomCmac, "", nCarLin, "OPERACIONES", sTitRp1, sTitRp2, sMoneda, Trim(nCntPag), psDescAge, sFchtra, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
    sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
    'sCad = sCad & "DESCRIPCION" & Space(24) & "NUM.REG      EFECTIVO       CHEQUES           O/P         TOTAL" & oImp.gPrnSaltoLinea
    sCad = sCad & "DESCRIPCION" & Space(24) & "NUM.REG      EFECTIVO       CHEQUES         O/P   NO EFECTIVO         TOTAL" & oImp.gPrnSaltoLinea 'FRHU 20150206 ERS048-2014
    sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & oImp.gPrnNegritaOFF
    nLinPag = 7
    If Not (rsMN.EOF And rsMN.BOF) Then
        If nLinPag > 58 Then
            sCad = sCad & Chr$(12)
            nCntPag = nCntPag + 1
            sCad = sCad & oImp.gPrnNegritaON
            sCad = sCad & oImpreCap.CabeRepoCaptac("", "", nCarLin, "OPERACIONES", sTitRp1, sTitRp2, sMoneda, Trim(nCntPag), gsNomAge, sFchtra, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
            sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
            'Cad = sCad & "DESCRIPCION" & Space(24) & "NUM.REG      EFECTIVO       CHEQUES           O/P         TOTAL" & oImp.gPrnSaltoLinea
            sCad = sCad & "DESCRIPCION" & Space(24) & "NUM.REG      EFECTIVO       CHEQUES         O/P   NO EFECTIVO         TOTAL" & oImp.gPrnSaltoLinea 'FRHU 20150206 ERS048-2014
            sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
            sCad = sCad & oImp.gPrnNegritaOFF
            nLinPag = 7
        End If
        rsMN.MoveFirst
        Do Until rsMN.EOF
              sNombre2 = Trim(rsMN(0))
              RSet sNumReg = Trim(rsMN(1))
              RSet sEfectivo = Trim(rsMN(2))
              RSet sCheques = Trim(rsMN(3))
              RSet sOrdPag = Trim(rsMN(4))
              'FRHU 20150206 ERS048-2014
              'RSet sTotFila = Trim(rsMN(5))
              RSet sNoEfectivo = Trim(rsMN(5))
              RSet sTotFila = Trim(rsMN(6))
              'FIN FRHU ERS048-2014
              If IsNumeric(Trim(sEfectivo)) Then nTotalEfectivo = nTotalEfectivo + CDbl(Trim(sEfectivo))
              If IsNumeric(Trim(sCheques)) Then nTotalCheque = nTotalCheque + CDbl(Trim(sCheques))
              If IsNumeric(Trim(sOrdPag)) Then nTotalOrdPag = nTotalOrdPag + CDbl(Trim(sOrdPag))
              If IsNumeric(Trim(sNoEfectivo)) Then nTotalNoEfectivo = nTotalNoEfectivo + CDbl(Trim(sNoEfectivo)) 'FRHU ERS048-2014
              If IsNumeric(Trim(sTotFila)) Then nTotalTotal = nTotalTotal + CDbl(Trim(sTotFila))
            
              If Trim(sEfectivo) = "" And Trim(UCase(sNombre2)) <> "TOTAL" Then
                  sCad = sCad & oImp.gPrnNegritaON
                  'sCad = sCad & sNombre2 & Space(5) & sNumReg & Space(2) & sEfectivo & Space(2) & sCheques & Space(2) & sOrdPag & Space(2) & sTotFila & oImp.gPrnSaltoLinea
                  sCad = sCad & sNombre2 & Space(5) & sNumReg & Space(2) & sEfectivo & Space(2) & sCheques & Space(2) & sOrdPag & Space(2) & sNoEfectivo & Space(2) & sTotFila & oImp.gPrnSaltoLinea 'FRHU 20150206 ERS048-2014
                  sCad = sCad & oImp.gPrnNegritaOFF
              ElseIf Trim(UCase(sNombre2)) = "TOTAL" Then
                  sCad = sCad & oImp.gPrnNegritaON
                  'sCad = sCad & sNombre2 & Space(2) & sNroMovimientos & Space(2) & sEfectivo & Space(2) & sCheques & Space(2) & sOrdPag & Space(2) & sTotFila & oImp.gPrnSaltoLinea
                  sCad = sCad & sNombre2 & Space(2) & sNroMovimientos & Space(2) & sEfectivo & Space(2) & sCheques & Space(2) & sOrdPag & Space(2) & sNoEfectivo & Space(2) & sTotFila & oImp.gPrnSaltoLinea 'FRHU 20150206 ERS048-2014
                  sCad = sCad & oImp.gPrnNegritaOFF
              Else
                  'sCad = sCad & sNombre2 & Space(5) & sNumReg & Space(2) & sEfectivo & Space(2) & sCheques & Space(2) & sOrdPag & Space(2) & sTotFila & oImp.gPrnSaltoLinea
                  sCad = sCad & sNombre2 & Space(5) & sNumReg & Space(2) & sEfectivo & Space(2) & sCheques & Space(2) & sOrdPag & Space(2) & sNoEfectivo & Space(2) & sTotFila & oImp.gPrnSaltoLinea 'FRHU 20150206 ERS048-2014
              End If
            rsMN.MoveNext
        Loop
    End If
    sCad = sCad & oImp.gPrnNegritaON
    sCad = sCad & String(nCarLin, "=") & oImp.gPrnSaltoLinea
    sNombre2 = "TOTAL"
    RSet sNumReg = ""
    RSet sEfectivo = Format$(nTotalEfectivo, "#,##0.00")
    RSet sCheques = Format$(nTotalCheque, "#,##0.00")
    RSet sOrdPag = Format$(nTotalOrdPag, "#,##0.00")
    RSet sNoEfectivo = Format$(nTotalNoEfectivo, "#,##0.00") 'FRHU 20150206 ERS048-2014
    RSet sTotFila = Format$(nTotalTotal, "#,##0.00")
    RSet sNroMovimientos = psNroMovMN
    'sCad = sCad & sNombre2 & Space(2) & sNroMovimientos & Space(2) & sEfectivo & Space(2) & sCheques & Space(2) & sOrdPag & Space(2) & sTotFila & oImp.gPrnSaltoLinea
    sCad = sCad & sNombre2 & Space(2) & sNroMovimientos & Space(2) & sEfectivo & Space(2) & sCheques & Space(2) & sOrdPag & Space(2) & sNoEfectivo & Space(2) & sTotFila & oImp.gPrnSaltoLinea 'FRHU 20150206 ERS048-2014
    sCad = sCad & oImp.gPrnSaltoLinea
    sCad = sCad & oImp.gPrnSaltoLinea
    sCad = sCad & "Total Ingresos = " & psTotalIngresosN & oImp.gPrnSaltoLinea
    sCad = sCad & "Total Egresos  = " & psTotalEgresosN & oImp.gPrnSaltoLinea
    sCad = sCad & "Total Habilita = " & psTotalHabilitacionN & oImp.gPrnSaltoLinea
    sCad = sCad & "Total Devoluci = " & psTotalDevolucionN & oImp.gPrnSaltoLinea
    sCad = sCad & "Total Dev.Bill = " & psTotalDevBilletajeN & oImp.gPrnSaltoLinea
    sCad = sCad & "Total Sob/Fal  = " & psTotalSobFaltN & oImp.gPrnSaltoLinea
    sCad = sCad & String(29, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & "TOTAL          = " & psTotalPlanillaN & oImp.gPrnSaltoLinea
     
    'If TxtBuscarUser.Text = "" Then
     
        sCad = sCad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
        sCad = sCad & "Saldo Anterior = " & psSaldoAnteriorN & oImp.gPrnSaltoLinea
        sCad = sCad & "Neto Movimient = " & psNetoMovimientosN & oImp.gPrnSaltoLinea
        '"   " & "Nro. Total Mov.=" & sNroMovimientos &
        sCad = sCad & String(29, "-") & oImp.gPrnSaltoLinea
        sCad = sCad & "SALDO FINAL    = " & psSaldoFinalN & oImp.gPrnSaltoLinea
        sCad = sCad & oImp.gPrnSaltoLinea
        
        '***PEAC 20090928
        'sCad = sCad & "DIFERENCIA     = " & psDiferenciaN & oImp.gPrnSaltoLinea
        If Not pbArqueo Then
            sCad = sCad & "DIFERENCIA     = " & psDiferenciaN & oImp.gPrnSaltoLinea
        End If
        '***FIN PEAC
    'End If
     
    sCad = sCad & "N° OPERACIONES = " & psNumOpeMN & oImp.gPrnSaltoLinea 'EJVG 20110809
    sCad = sCad & oImp.gPrnNegritaOFF
    
    'Moneda Extranjera
    sCad = sCad & Chr$(12)
    nTotalEfectivo = 0
    nTotalCheque = 0
    nTotalOrdPag = 0
    nTotalNoEfectivo = 0 'FRHU 20150206 ERS048-2014
    nTotalTotal = 0
    sMoneda = "DOLARES"
    sCad = sCad & oImp.gPrnNegritaON
    sCad = sCad & oImpreCap.CabeRepoCaptac("", "", nCarLin, "OPERACIONES", sTitRp1, sTitRp2, sMoneda, Trim(nCntPag), gsNomAge, sFchtra, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
    sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
    'sCad = sCad & "DESCRIPCION" & Space(24) & "NUM.REG      EFECTIVO       CHEQUES           O/P         TOTAL" & oImp.gPrnSaltoLinea
    sCad = sCad & "DESCRIPCION" & Space(24) & "NUM.REG      EFECTIVO       CHEQUES         O/P     NO EFECTIVO         TOTAL" & oImp.gPrnSaltoLinea 'FRHU 20150206 ERS048-2014
    sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & oImp.gPrnNegritaOFF
    nLinPag = 7
    If Not (rsME.EOF And rsME.BOF) Then
        If nLinPag > 58 Then
            sCad = sCad & Chr$(12)
            nCntPag = nCntPag + 1
            sCad = sCad & oImp.gPrnNegritaON
            sCad = sCad & oImpreCap.CabeRepoCaptac("", "", nCarLin, "OPERACIONES", sTitRp1, sTitRp2, sMoneda, Trim(nCntPag), gsNomAge, sFchtra, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
            sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
            'sCad = sCad & "DESCRIPCION" & Space(24) & "NUM.REG      EFECTIVO       CHEQUES           O/P         TOTAL" & oImp.gPrnSaltoLinea
            sCad = sCad & "DESCRIPCION" & Space(24) & "NUM.REG      EFECTIVO       CHEQUES         O/P     NO EFECTIVO         TOTAL" & oImp.gPrnSaltoLinea 'FRHU 20150206 ERS048-2014
            sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
            sCad = sCad & oImp.gPrnNegritaOFF
            nLinPag = 7
        End If
        rsME.MoveFirst
        Do Until rsME.EOF
             sNombre = Trim(rsME(0))
             RSet sNumReg = Trim(rsME(1))
             RSet sEfectivo = Trim(rsME(2))
             RSet sCheques = Trim(rsME(3))
             RSet sOrdPag = Trim(rsME(4))
             'FRHU 20150206 ERS048-2014
             'RSet sTotFila = Trim(rsME(5))
             RSet sNoEfectivo = Trim(rsME(5))
             RSet sTotFila = Trim(rsME(6))
             'FIN FRHU 20150206
             If IsNumeric(Trim(sEfectivo)) Then nTotalEfectivo = nTotalEfectivo + CDbl(Trim(sEfectivo))
             If IsNumeric(Trim(sCheques)) Then nTotalCheque = nTotalCheque + CDbl(Trim(sCheques))
             If IsNumeric(Trim(sOrdPag)) Then nTotalOrdPag = nTotalOrdPag + CDbl(Trim(sOrdPag))
             If IsNumeric(Trim(sNoEfectivo)) Then nTotalNoEfectivo = nTotalNoEfectivo + CDbl(Trim(sNoEfectivo)) 'FRHU 20150206 ERS048-2014
             If IsNumeric(Trim(sTotFila)) Then nTotalTotal = nTotalTotal + CDbl(Trim(sTotFila))
            
             If Trim(sEfectivo) = "" Then
                 sCad = sCad & oImp.gPrnNegritaON
                 'sCad = sCad & sNombre & Space(2) & sNumReg & Space(2) & sEfectivo & Space(2) & sCheques & Space(2) & sOrdPag & Space(2) & sTotFila & oImp.gPrnSaltoLinea
                 sCad = sCad & sNombre & Space(2) & sNumReg & Space(2) & sEfectivo & Space(2) & sCheques & Space(2) & sOrdPag & Space(2) & sNoEfectivo & Space(2) & sTotFila & oImp.gPrnSaltoLinea
                 sCad = sCad & oImp.gPrnNegritaOFF
             Else
                 'sCad = sCad & sNombre & Space(2) & sNumReg & Space(2) & sEfectivo & Space(2) & sCheques & Space(2) & sOrdPag & Space(2) & sTotFila & oImp.gPrnSaltoLinea
                 sCad = sCad & sNombre & Space(2) & sNumReg & Space(2) & sEfectivo & Space(2) & sCheques & Space(2) & sOrdPag & Space(2) & sNoEfectivo & Space(2) & sTotFila & oImp.gPrnSaltoLinea 'FRHU 20150206 ERS048-2014
             End If
            rsME.MoveNext
        Loop
    End If
    
    sCad = sCad & oImp.gPrnNegritaON
    sCad = sCad & String(nCarLin, "=") & oImp.gPrnSaltoLinea
    sNombre2 = "TOTAL"
    RSet sNumReg = ""
    RSet sEfectivo = Format$(nTotalEfectivo, "#,##0.00")
    RSet sCheques = Format$(nTotalCheque, "#,##0.00")
    RSet sOrdPag = Format$(nTotalOrdPag, "#,##0.00")
    RSet sNoEfectivo = Format$(nTotalNoEfectivo, "#,##0.00") 'FRHU 20150206 ERS048-2014
    RSet sTotFila = Format$(nTotalTotal, "#,##0.00")
    RSet sNroMovimientos = psNroMovME
    'sCad = sCad & sNombre2 & Space(2) & sNroMovimientos & Space(2) & sEfectivo & Space(2) & sCheques & Space(2) & sOrdPag & Space(2) & sTotFila & oImp.gPrnSaltoLinea
    sCad = sCad & sNombre2 & Space(2) & sNroMovimientos & Space(2) & sEfectivo & Space(2) & sCheques & Space(2) & sOrdPag & Space(2) & sNoEfectivo & Space(2) & sTotFila & oImp.gPrnSaltoLinea
    sCad = sCad & oImp.gPrnSaltoLinea
    sCad = sCad & oImp.gPrnSaltoLinea
    sCad = sCad & "Total Ingresos = " & psTotalIngresosE & oImp.gPrnSaltoLinea
    sCad = sCad & "Total Egresos  = " & psTotalEgresosE & oImp.gPrnSaltoLinea
    sCad = sCad & "Total Habilita = " & psTotalHabilitacionE & oImp.gPrnSaltoLinea
    sCad = sCad & "Total Devoluci = " & psTotalDevolucionE & oImp.gPrnSaltoLinea
    sCad = sCad & "Total Dev.Bill = " & psTotalDevBilletajeE & oImp.gPrnSaltoLinea
    sCad = sCad & "Total Sob/Fal  = " & psTotalSobFaltE & oImp.gPrnSaltoLinea
    sCad = sCad & String(29, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & "TOTAL          = " & psTotalPlanillaE & oImp.gPrnSaltoLinea
    'If TxtBuscarUser.Text = "" Then
       
        sCad = sCad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
        sCad = sCad & "Saldo Anterior = " & psSaldoAnteriorE & oImp.gPrnSaltoLinea
        sCad = sCad & "Neto Movimient = " & psNetoMovimientosE & oImp.gPrnSaltoLinea
        '"   " & "Nro. Total Mov.=" & sNroMovimientos &
        sCad = sCad & String(29, "-") & oImp.gPrnSaltoLinea
        sCad = sCad & "SALDO FINAL    = " & psSaldoFinalE & oImp.gPrnSaltoLinea
        sCad = sCad & oImp.gPrnSaltoLinea
        
        '***PEAC 20090928
        'sCad = sCad & "DIFERENCIA     = " & psDiferenciaE & oImp.gPrnSaltoLinea
        If Not pbArqueo Then
            sCad = sCad & "DIFERENCIA     = " & psDiferenciaE & oImp.gPrnSaltoLinea
        End If
        '***FIN PEAC
        
    'End If
    
    sCad = sCad & "N° OPERACIONES = " & psNumOpeME & oImp.gPrnSaltoLinea 'EJVG 20110809
    
'*** RIRO 20130702 SEGUN TI-ERS083-2013 **********************************************************************************************
If Not rsEgresos Is Nothing Then
    If Not rsEgresos.EOF And Not rsEgresos.BOF Then
        Dim nEspacio As Integer
        nEspacio = 20
        sCad = sCad & oImp.gPrnSaltoLinea
        sCad = sCad & "Nº Operaciones con la opción Otros Egresos En Efectivo: " & Format(rsEgresos("Registros"), "##00") & oImp.gPrnSaltoLinea
        sCad = sCad & String(100, "-") & oImp.gPrnSaltoLinea
        sCad = sCad & Left("USUARIO" & Space(20), nEspacio / 2) _
                    & Left("GLOSA" & Space(20), nEspacio) _
                    & Left("NUM. OPERACION" & Space(20), nEspacio) _
                    & Left("HORA" & Space(20), nEspacio) _
                    & Left("MONEDA" & Space(20), nEspacio) _
                    & Left("MONTO" & Space(20), nEspacio) & oImp.gPrnSaltoLinea
        sCad = sCad & String(100, "-") & oImp.gPrnSaltoLinea
        Do Until rsEgresos.EOF
            sCad = sCad & Left(Space(1) & Trim(Replace(rsEgresos("CUSER"), vbNewLine, " ")) & Space(20), nEspacio / 2) _
                        & Left(Trim(Replace(rsEgresos("GLOSA"), vbNewLine, " ")) & Space(20), nEspacio) _
                        & Left(Space(5) & Trim(Replace(rsEgresos("OPERACION"), vbNewLine, " ")) & Space(20), nEspacio) _
                        & Left(Trim(Replace(rsEgresos("HORA"), vbNewLine, " ")) & Space(20), nEspacio) _
                        & Left(Trim(Replace(rsEgresos("MONEDA"), vbNewLine, " ")) & Space(20), nEspacio) _
                        & Left(Trim(Replace(Format(rsEgresos("MONTO"), "##0.00"), vbNewLine, " ")) & Space(20), nEspacio) & oImp.gPrnSaltoLinea
            rsEgresos.MoveNext
        Loop
        rsEgresos.MoveFirst
        sCad = sCad & String(100, "=") & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
        '''sCad = sCad & "TOTAL SOLES    : " & Format(rsEgresos("Soles"), "#,##0.00") & oImp.gPrnSaltoLinea 'marg ers044-2016
        sCad = sCad & "TOTAL " & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & "    : " & Format(rsEgresos("Soles"), "#,##0.00") & oImp.gPrnSaltoLinea 'marg ers044-2016
        sCad = sCad & "TOTAL DOLARES  : " & Format(rsEgresos("Dolares"), "#,##0.00") & oImp.gPrnSaltoLinea
    End If
    
End If
'FIN RIRO ****************************************************************************************************************************
    
    sCad = sCad & oImp.gPrnNegritaOFF

  ImpCajeroIngEgre = sCad
Set oImpreCap = Nothing
End Function

Public Function ImpPrecuadre(ByVal gsCodAge As String, ByVal gsNomAge As String, _
    ByVal psDifMN As String, ByVal psDifME As String, ByVal gdFecSis As Date, _
    Optional ByVal pbValidaSobFal As Boolean = False) As String
Dim oCaj As COMNCajaGeneral.NCOMCajero  'nCajero
Dim rs As adodb.Recordset
Dim sCad As String
Dim sFecha As String * 10
Dim sOperacion As String * 30, sPersona As String * 30
Dim sUsuario As String * 4, sMoneda As String * 3, sUsuarioDest As String * 4
Dim sMonto As String * 12
Dim nCarLin As Long
Dim sTitRp1 As String, sTitRp2 As String
Dim oImpreCap As New COMNCaptaGenerales.NCOMCaptaImpresion
 
'Verifica que no hayan confirmaciones de habilitaciones de cajeros
sOperacion = gOpeBoveAgeHabCajero & "','" & gOpeHabCajTransfEfectCajeros
Set oCaj = New COMNCajaGeneral.NCOMCajero
    Set rs = oCaj.GetMovHabBovAgeCajero(sOperacion, "", gsCodAge, gdFecSis, gdFecSis)
Set oCaj = Nothing
If Not (rs.EOF And rs.BOF) Then
    nCarLin = 73
    sTitRp1 = "HABILITACIONES A CAJERO AUN NO CONFIRMADAS"
    sTitRp2 = ""
    sCad = sCad & oImpreCap.CabeRepoCaptac("", "", nCarLin, "SECCION OPERACIONES", sTitRp1, sTitRp2, "", "   1", Trim(gsNomAge), gdFecSis, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
    sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & "FECHA       OPERACION                       ORIG  DEST              MONTO" & oImp.gPrnSaltoLinea
    sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
    Do While Not rs.EOF
        sFecha = Left(rs("Fecha"), 10)
        sOperacion = Trim(rs("cOpeDesc"))
        sUsuarioDest = Trim(rs("cUsuDest"))
        sUsuario = Trim(rs("cUsuOrig"))
        sMoneda = rs("cSimbolo")
        RSet sMonto = Format$(rs("nMovImporte"), "#,##0.00")
        sCad = sCad & sFecha & Space(2) & sOperacion & Space(2) & sUsuario & Space(2) & sUsuarioDest & Space(2) & sMoneda & Space(2) & sMonto & oImp.gPrnSaltoLinea
        rs.MoveNext
    Loop
End If

'Verificamos que Boveda no Tenga confirmaciones de devoluciones
sOperacion = gOpeHabCajDevABove
Set oCaj = New COMNCajaGeneral.NCOMCajero
Set rs = oCaj.GetMovHabCajDevABove(sOperacion, "", gsCodAge, gdFecSis, gdFecSis)
Set oCaj = Nothing
If Not (rs.EOF And rs.BOF) Then
    nCarLin = 99
    sTitRp1 = "HABILITACIONES A BOVEDA AUN NO CONFIRMADAS"
    sTitRp2 = ""
    If sCad <> "" Then sCad = sCad & Chr$(12)
    sCad = sCad & oImpreCap.CabeRepoCaptac("", "", nCarLin, "SECCION OPERACIONES", sTitRp1, sTitRp2, "", "   1", Trim(gsNomAge), gdFecSis, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
    sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & "FECHA       OPERACION                       USU   PERSONA                                     MONTO" & oImp.gPrnSaltoLinea
    sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
    Do While Not rs.EOF
        sFecha = Left(rs("Fecha"), 10)
        sOperacion = Trim(rs("cOpeDesc"))
        sPersona = Trim(rs("cPersNombre"))
        sUsuario = Trim(rs("cUser"))
        sMoneda = rs("cSimbolo")
        RSet sMonto = Format$(rs("nMovImporte"), "#,##0.00")
        sCad = sCad & sFecha & Space(2) & sOperacion & Space(2) & sUsuario & Space(2) & sPersona & Space(2) & sMoneda & Space(2) & sMonto & oImp.gPrnSaltoLinea
        rs.MoveNext
    Loop
End If

'Verifica Usuarios que hayan hecho movimientos en efectivo y aun no han hecho devolucion por billetaje
Dim bCabecera As Boolean
bCabecera = False
Set oCaj = New COMNCajaGeneral.NCOMCajero
Set rs = oCaj.GetUsuariosMovEfectivo(gsCodAge, gdFecSis)
If Not (rs.EOF And rs.BOF) Then
    Do While Not rs.EOF
        If Not oCaj.YaRealizoDevBilletaje(rs("Usuario"), gdFecSis, gsCodAge) Then
            If Not bCabecera Then
                nCarLin = 73
                sTitRp1 = "USUARIOS QUE NO HAN REALIZADO DEVOLUCION BILLETAJE"
                sTitRp2 = ""
                If sCad <> "" Then sCad = sCad & Chr$(12)
                sCad = sCad & oImpreCap.CabeRepoCaptac("", "", nCarLin, "SECCION OPERACIONES", sTitRp1, sTitRp2, "", "   1", Trim(gsNomAge), gdFecSis, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
                sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
                sCad = sCad & "USU   NOMBRE" & oImp.gPrnSaltoLinea
                sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
                bCabecera = True
            End If
            sUsuario = Trim(rs("Usuario"))
            sCad = sCad & sUsuario & Space(2) & Trim(rs("cPersNombre")) & oImp.gPrnSaltoLinea
        End If
        rs.MoveNext
    Loop
End If
Set oCaj = Nothing

'Verifica Diferencias en la Planilla
Dim nDifMN As Double, nDifME As Double
If IsNumeric(psDifMN) And IsNumeric(psDifME) Then
    nDifMN = CDbl(psDifMN)
    nDifME = CDbl(psDifME)
Else
    nDifMN = 0
    nDifME = 0
End If
If nDifMN <> 0 Or nDifME <> 0 Then
    nCarLin = 73
    sTitRp1 = "DIFERENCIAS EN RESUMEN DE INGRESOS Y EGRESOS CONSOLIDADO"
    sTitRp2 = ""
    If sCad <> "" Then sCad = sCad & Chr$(12)
    sCad = sCad & oImpreCap.CabeRepoCaptac("", "", nCarLin, "SECCION OPERACIONES", sTitRp1, sTitRp2, "", "   1", Trim(gsNomAge), gdFecSis, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
    sCad = sCad & oImp.gPrnSaltoLinea
    If nDifMN <> 0 Then
        RSet sMonto = Format$(nDifMN, "#,##0.00")
        '''sCad = sCad & "DIFERENCIA MONEDA NACIONAL S/.   = " & sMonto & oImp.gPrnSaltoLinea 'marg ers044-2016
        sCad = sCad & "DIFERENCIA MONEDA NACIONAL " & gcPEN.gcPEN_SIMBOLO & "   = " & sMonto & oImp.gPrnSaltoLinea   'marg ers044-2016
    End If
    If nDifME <> 0 Then
        RSet sMonto = Format$(nDifME, "#,##0.00")
        sCad = sCad & "DIFERENCIA MONEDA EXTRANJERA US$ = " & sMonto & oImp.gPrnSaltoLinea
    End If
    
End If

''Verificamos operaciones
If pbValidaSobFal Then
    Set oCaj = New COMNCajaGeneral.NCOMCajero
    Set rs = oCaj.GetSobFalMayUNO(gdFecSis, gsCodAge)
    If Not (rs.EOF And rs.BOF) Then
        'Do While Not rs.EOF
         nCarLin = 99
        sTitRp1 = "SOBRANTES Y FALTANTES MAYORES A UNO"
        sTitRp2 = ""
        If sCad <> "" Then sCad = sCad & oImp.gPrnSaltoPagina
        sCad = sCad & oImpreCap.CabeRepoCaptac("", "", nCarLin, "SECCION OPERACIONES", sTitRp1, sTitRp2, "", "   1", Trim(gsNomAge), gdFecSis, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
        sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
        sCad = sCad & "USUARIO       OPERACION                       MONTO  MONEDA" & oImp.gPrnSaltoLinea
        sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
        Do While Not rs.EOF
            sUsuario = Trim(rs("cUser"))
            sOperacion = Trim(rs("cOpeDesc"))
            sMoneda = rs("cSimbolo")
            sMonto = Format$(rs("nMovImporte"), "#,##0.00")
            sCad = sCad & sFecha & Space(2) & sUsuario & Space(2) & sOperacion & Space(2) & sMonto & Space(2) & sMoneda & oImp.gPrnSaltoLinea
            rs.MoveNext
        Loop
    End If
End If
Set oCaj = Nothing
ImpPrecuadre = sCad
Set oImpreCap = Nothing
End Function

Public Function ImpreRegularizacion(ByVal gsNomAge As String, ByVal gcEmpresa As String, ByVal gdFecSis As Date, Optional ByVal pnOpeCod As Long) As String
    Dim lsCodUsu As String * 6
    Dim lsNomUsu As String * 35
    Dim lsMonto As String * 13
    Dim lsTipo As String * 8
    Dim lsFecha As String * 15
    Dim lsNroMov As String * 10
    Dim lsMoneda As String * 10
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lnI As Integer
    Dim sql As String
    Dim rs As adodb.Recordset
    Set rs = New adodb.Recordset
    Dim lsMonedaAnt As String
    Dim lsOpeAnt As String
    Dim lnAcum As Currency
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta

    Dim oFunI As COMFunciones.FCOMImpresion
    Set oFunI = New COMFunciones.FCOMImpresion
    
    oCon.AbreConexion
    
    'MAVM 20120328 ***
    'Sql = " Select Right(M.cMovNro,4) cCodusu, ISNULL(PE.cPersNombre,'') cNomUsu, MSF.nMoneda  cCodCta, MSF.nMovImporte nMontran, M.nMovNro cNroTran, dbo.GetFechaMov(M.cMovNro,103) dFecTran, M.cOpeCod cCodOpe " _
    '    & " From Mov M " _
    '    & " Inner Join MovOpeVarias MSF On M.nMovNro = MSF.nMovNro " _
    '    & " Left  Join RRHH RH On RH.cUser = Right(M.cMovNro,4) " _
    '    & " Left  Join Persona PE On PE.cPersCod = RH.cPersCod " _
    '    & " Where  ((M.cOpeCod = '" & gOpeHabCajRegSobrante & "' And Abs(MSF.nMovImporte) > 5) " _
    '    & "     Or      (M.cOpeCod = '" & gOpeHabCajRegFaltante & "' And Abs(MSF.nMovImporte) > 1)) And M.nMovFlag = 0 " _
    '    & " And Not Exists (Select MR.nMovNro From Mov MR " _
    '    & "         Inner Join MovRef MRR On MR.nMovNro = MRR.nMovNro " _
    '    & "         Where MR.nMovFlag = " & MovFlag.gMovFlagVigente & " And MRR.nMovNroRef = M.nMovNro) " _
    '    & " Order By  M.cMovNro Desc"
        
        sql = " Select Right(M.cMovNro,4) cCodusu, ISNULL(PE.cPersNombre,'') cNomUsu, MSF.nMoneda  cCodCta, MSF.nMovImporte nMontran, M.nMovNro cNroTran, dbo.GetFechaMov(M.cMovNro,103) dFecTran, M.cOpeCod cCodOpe " _
        & " From Mov M " _
        & " Inner Join MovOpeVarias MSF On M.nMovNro = MSF.nMovNro " _
        & " Left  Join RRHH RH On RH.cUser = Right(M.cMovNro,4) " _
        & " Left  Join Persona PE On PE.cPersCod = RH.cPersCod " _
        & " Where M.cOpeCod in ('901019','901020','901041','901042') " _
        & " And M.nMovFlag = 0 " _
        & " And Not Exists (Select MR.nMovNro From Mov MR " _
        & "         Inner Join MovRef MRR On MR.nMovNro = MRR.nMovNro " _
        & "         Where MR.nMovFlag = " & MovFlag.gMovFlagVigente & " And MRR.nMovNroRef = M.nMovNro) " _
        & " And (left(M.cMovNro,4) >= '2010') " _
        & " Order By  M.cMovNro Desc"
    '***
    
    Set rs = oCon.CargaRecordSet(sql)
    lsCadena = ""
    
    If Not (rs.EOF And rs.BOF) Then
        lsCadena = lsCadena & oFunI.CabeceraPagina("Pendientes de Caja (Sobrantes/Faltantes) Agencia : " & gsNomAge, lnPagina, lnItem, gsNomAge, gcEmpresa, gdFecSis, rs!cCodCta)
        lsCadena = lsCadena & oFunI.Encabezado("Usurio;6; ;2;Nombre;10; ;27;Tipo;7; ;5;Monto;10;Moneda;10; ;5;Fecha;10; ;5;NroRef;10; ;5;", lnItem)
        
        While Not rs.EOF
            
            If (lsMonedaAnt <> Mid(rs!cCodCta, 6, 1) Or lsOpeAnt <> rs!cCodOpe) And lsMonedaAnt <> "" Then
                lsCodUsu = ""
                lsNomUsu = ""
                lsTipo = ""
                lsMonto = ""
                lsMoneda = ""
                lsFecha = ""
                lsNroMov = ""
                RSet lsMonto = Format(lnAcum, "#,##0.00")
                lnAcum = 0
                lsCadena = lsCadena & String(Len(lsCodUsu & "  " & lsNomUsu & "  " & lsTipo & "  " & lsMonto & "  " & lsMoneda & "  " & lsFecha & "  " & lsNroMov), "=") & Chr(10)
                lsCadena = lsCadena & lsCodUsu & "  " & lsNomUsu & "  " & lsTipo & "  " & lsMonto & "  " & lsMoneda & "  " & lsFecha & "  " & lsNroMov & Chr(10)
                lsCadena = lsCadena & Chr(12)
                lsCadena = lsCadena & oFunI.CabeceraPagina("Pendientes de Caja (Sobrantes/Faltantes) Agencia : " & gsNomAge, lnPagina, lnItem, gsNomAge, gcEmpresa, gdFecSis, Mid(rs!cCodCta, 9, 1))
                lsCadena = lsCadena & oFunI.Encabezado("Usurio;6; ;2;Nombre;10; ;27;Tipo;7; ;5;Monto;10;Moneda;10; ;5;Fecha;10; ;5;NroRef;10; ;5;", lnItem)
            End If
            
            lsCodUsu = rs!cCodUsu
            lsNomUsu = rs!cNomusu
            
            If pnOpeCod = gOtrOpePagoFaltante Then
                If rs!nMonTran < 0 Then
                    RSet lsMonto = Format(Abs(rs!nMonTran), "#,##0.00")
                    lnAcum = lnAcum + Abs(rs!nMonTran)
                
                    lsTipo = "FALTANTE"
                    lsFecha = Format(rs!dFecTran, "dd/mm/yyyy")
                    lsNroMov = rs!cNroTran
                    If rs!cCodCta <> "2" Then
                        '''lsMoneda = "Soles" 'marg ers044-2016
                        lsMoneda = StrConv(gcPEN.gcPEN_PLURAL, vbProperCase) 'marg ers044-2016
                    Else
                        lsMoneda = "Dolares"
                    End If
                    lsCadena = lsCadena & lsCodUsu & "  " & lsNomUsu & "  " & lsTipo & "  " & lsMonto & "  " & lsMoneda & "  " & lsFecha & "  " & lsNroMov & Chr(10)
                    lnItem = lnItem + 1
                    lsMonedaAnt = Mid(rs!cCodCta, 6, 1)
                    lsOpeAnt = rs!cCodOpe
                End If
            Else
                If rs!nMonTran > 0 Then
                    RSet lsMonto = Format(Abs(rs!nMonTran), "#,##0.00")
                    lnAcum = lnAcum + Abs(rs!nMonTran)
                    
                    lsTipo = "SOBRANTE"
                    lsFecha = Format(rs!dFecTran, "dd/mm/yyyy")
                    lsNroMov = rs!cNroTran
                    If rs!cCodCta <> "2" Then
                        '''lsMoneda = "Soles" 'marg ers044-2016
                        lsMoneda = StrConv(gcPEN.gcPEN_PLURAL, vbProperCase) 'marg ers044-2016
                    Else
                        lsMoneda = "Dolares"
                    End If
                    lsCadena = lsCadena & lsCodUsu & "  " & lsNomUsu & "  " & lsTipo & "  " & lsMonto & "  " & lsMoneda & "  " & lsFecha & "  " & lsNroMov & Chr(10)
                    lnItem = lnItem + 1
                    lsMonedaAnt = Mid(rs!cCodCta, 6, 1)
                    lsOpeAnt = rs!cCodOpe
                End If
            End If
            
            rs.MoveNext
        Wend
    End If
    
    
    lsCodUsu = ""
    lsNomUsu = ""
    lsTipo = ""
    lsMonto = ""
    lsMoneda = ""
    lsFecha = ""
    lsNroMov = ""
    RSet lsMonto = Format(lnAcum, "#,##0.00")
    lnAcum = 0
    lsCadena = lsCadena & String(Len(lsCodUsu & "  " & lsNomUsu & "  " & lsTipo & "  " & lsMonto & "  " & lsMoneda & "  " & lsFecha & "  " & lsNroMov), "=") & Chr(10)
    lsCadena = lsCadena & lsCodUsu & "  " & lsNomUsu & "  " & lsTipo & "  " & lsMonto & "  " & lsMoneda & "  " & lsFecha & "  " & lsNroMov & Chr(10)
    
    ImpreRegularizacion = lsCadena
    
    oCon.CierraConexion
End Function

Public Function ObtenerRegularizacion(ByVal pnMonSobrante As Double, ByVal pnMonFaltante As Double) As adodb.Recordset
     Dim oCon As COMConecta.DCOMConecta
     Dim sql As String
     Dim rs As adodb.Recordset
     Set oCon = New COMConecta.DCOMConecta
 On Error GoTo xError
     oCon.AbreConexion
     
    '*** PEAC 20090924
'     Sql = " Select Right(M.cMovNro,4) cCodusu, PE.cPersNombre cNomUsu, MSF.nMoneda  cCodCta, MSF.nMovImporte nMontran, M.nMovNro cNroTran, dbo.GetFechaMov(M.cMovNro,103) dFecTran, M.cOpeCod cCodOpe " _
'        & " From Mov M INNER JOIN MovOpeVarias MSF On M.nMovNro = MSF.nMovNro " _
'        & " LEFT JOIN RRHH RH On RH.cUser = Right(M.cMovNro,4) LEFT JOIN Persona PE On PE.cPersCod = RH.cPersCod " _
'        & " Where  ((M.cOpeCod = '" & gOpeHabCajRegSobrante & "' And Abs(MSF.nMovImporte) > " & pnMonSobrante & ") " _
'        & "     Or      (M.cOpeCod = '" & gOpeHabCajRegFaltante & "' And Abs(MSF.nMovImporte) > " & pnMonFaltante & ")) And M.nMovFlag = 0 " _
'        & " And Not Exists (Select MR.nMovNro From Mov MR " _
'        & "         Inner Join MovRef MRR On MR.nMovNro = MRR.nMovNro " _
'        & "         Where MR.nMovFlag = " & MovFlag.gMovFlagVigente & " And MRR.nMovNroRef = M.nMovNro) " _
'        & " Order By  M.cMovNro Desc"
        
    sql = " exec stp_sel_ObtenerRegularizacionSobranteFaltante "
        
    Set rs = oCon.CargaRecordSet(sql)
    Set ObtenerRegularizacion = rs
    Exit Function
xError:
    Set ObtenerRegularizacion = Nothing
    oCon.CierraConexion
End Function

'**DAOR 20080125, Obtener registros para validación de billetaje
Public Function ObtenerValidacionRegistroEfectivo(ByVal psCodUser As String, ByVal psCodAge As String, ByVal psFecha As String, ByVal psOpeCod As String) As adodb.Recordset
Dim loCon As COMConecta.DCOMConecta
Dim lsSql As String
Dim lRs As adodb.Recordset
On Error GoTo ObtenerValidacionRegistroEfectivoErr
    Set loCon = New COMConecta.DCOMConecta
        loCon.AbreConexion
    lsSql = " exec stp_sel_ValidacionRegistroEfectivo '" & psCodUser & "','" & psCodAge & "','" & psFecha & "','" & psOpeCod & "'"
    Set lRs = loCon.CargaRecordSet(lsSql)
    Set ObtenerValidacionRegistroEfectivo = lRs
    loCon.CierraConexion
    Exit Function
ObtenerValidacionRegistroEfectivoErr:
    Set ObtenerValidacionRegistroEfectivo = Nothing
    loCon.CierraConexion
End Function

'**DAOR 20080125, Graba datos de resumen del registro de efectivo
Public Function GrabaRegistroEfectivoResumen(ByVal pnMovNro As Long, ByVal psUser As String, ByVal pnNumRegEfec As Integer, ByVal pnIngresosMN As Currency, ByVal pnEgresosMN As Currency, ByVal pnEfectivoMN As Currency, ByVal pnSobrFaltMN As Currency, _
    ByVal pnIngresosME As Currency, ByVal pnEgresosME As Currency, ByVal pnEfectivoME As Currency, ByVal pnSobrFaltME As Currency) As Integer
Dim loMov As COMDMov.DCOMMov
Dim lbTrans As Boolean
On Error GoTo GrabaRegistroEfectivoResumenErr
    Set loMov = New COMDMov.DCOMMov
    loMov.BeginTrans
    lbTrans = True
    GrabaRegistroEfectivoResumen = 1
    loMov.InsertaMovRegEfectivo pnMovNro, psUser, pnNumRegEfec, pnIngresosMN, pnEgresosMN, pnEfectivoMN, pnSobrFaltMN, pnIngresosME, pnEgresosME, pnEfectivoME, pnSobrFaltME
    loMov.CommitTrans
    lbTrans = False
    GrabaRegistroEfectivoResumen = 0
    Set loMov = Nothing
    Exit Function
GrabaRegistroEfectivoResumenErr:
    If lbTrans Then
        loMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'MADM 20110127
Public Function GrabaRegistroEfectivoResumenPreCuadre(ByVal pnMovNro As Long, ByVal psUser As String, ByVal pnNumRegEfec As Integer, ByVal pnIngresosMN As Currency, ByVal pnEgresosMN As Currency, ByVal pnEfectivoMN As Currency, ByVal pnSobrFaltMN As Currency, _
    ByVal pnIngresosME As Currency, ByVal pnEgresosME As Currency, ByVal pnEfectivoME As Currency, ByVal pnSobrFaltME As Currency) As Integer
Dim loMov As COMDMov.DCOMMov
Dim lbTrans As Boolean
On Error GoTo GrabaRegistroEfectivoResumenCorteErr
    Set loMov = New COMDMov.DCOMMov
    loMov.BeginTrans
    lbTrans = True
    GrabaRegistroEfectivoResumenPreCuadre = 1
    loMov.InsertaMovRegEfectivoPreCuadre pnMovNro, psUser, pnNumRegEfec, pnIngresosMN, pnEgresosMN, pnEfectivoMN, pnSobrFaltMN, pnIngresosME, pnEgresosME, pnEfectivoME, pnSobrFaltME
    loMov.CommitTrans
    lbTrans = False
    GrabaRegistroEfectivoResumenPreCuadre = 0
    Set loMov = Nothing
    Exit Function
GrabaRegistroEfectivoResumenCorteErr:
    If lbTrans Then
        loMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'END MADM


'**DAOR 20080125, Obtener registros para validación de billetaje
Public Function ObtenerReporteRegistrosEfectivoPorUsuario(ByVal psFechaI As String, ByVal psFechaF As String, ByVal psListaAgencias As String) As adodb.Recordset
Dim loCon As COMConecta.DCOMConecta
Dim lsSql As String
On Error GoTo ObtenerReporteRegistrosEfectivoPorUsuarioErr
    Set loCon = New COMConecta.DCOMConecta
        loCon.AbreConexion
    lsSql = " exec stp_sel_ReporteRegistrosDeEfectivo '" & psFechaI & "','" & psFechaF & "','" & psListaAgencias & "'"
    Set ObtenerReporteRegistrosEfectivoPorUsuario = loCon.CargaRecordSet(lsSql)
    loCon.CierraConexion
    Exit Function
ObtenerReporteRegistrosEfectivoPorUsuarioErr:
    Set ObtenerReporteRegistrosEfectivoPorUsuario = Nothing
    loCon.CierraConexion
End Function

'**DAOR 20080204
Public Function ObtenerRegistrosDeEfectivo(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                    ByVal psCodUser As String, psCodAge As String) As adodb.Recordset
Dim loConecta As COMConecta.DCOMConecta
Dim lsSql As String
    
    Set loConecta = New COMConecta.DCOMConecta
    If loConecta.AbreConexion = False Then Exit Function

    lsSql = "exec stp_sel_RegistrosDeEfectivo '" & psOpeCod & "','" & Format(pdDesde, "yyyymmdd") & "','" & Format(pdHasta, "yyyymmdd") & "','" & psCodUser & "','" & psCodAge & "'"
 
    Set ObtenerRegistrosDeEfectivo = loConecta.CargaRecordSet(lsSql)
    loConecta.CierraConexion
    Set loConecta = Nothing
End Function

Public Function obtenerGruposOpeITF() As Recordset
    Dim loConecta As COMConecta.DCOMConecta
    Dim lsSql As String
    
    Set loConecta = New COMConecta.DCOMConecta
    If loConecta.AbreConexion = False Then Exit Function

    lsSql = "exec stp_sel_ObtenerGruposOpeITF"
 
    Set obtenerGruposOpeITF = loConecta.CargaRecordSet(lsSql)
    loConecta.CierraConexion
    Set loConecta = Nothing
End Function

Private Sub Class_Terminate()
Set oImp = Nothing
End Sub

''MADM 20120328 - 20111227 - Comentado x MADM - aun no se firma el acta de conformidad
Public Function ObtenerRegularizacionSobranteCastigo(ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal pnMoneda As Integer, ByVal ind As Integer) As adodb.Recordset
     Dim oCon As COMConecta.DCOMConecta
     Dim sql As String
     Dim rs As adodb.Recordset
     Set oCon = New COMConecta.DCOMConecta
 On Error GoTo xError
     oCon.AbreConexion

    sql = " exec stp_sel_ObtenerRegularizacionSobranteFaltantexCastigo " & pnMoneda & ", '" & Format(pdDesde, "yyyymmdd") & "','" & Format(pdHasta, "yyyymmdd") & "', " & ind & " "
    
    Set rs = oCon.CargaRecordSet(sql)
    Set ObtenerRegularizacionSobranteCastigo = rs
    Exit Function
xError:
    Set ObtenerRegularizacionSobranteCastigo = Nothing
    oCon.CierraConexion
End Function

'** Juez 20120815 ********************************************************************
Public Function GetOpeIngEgreCaptacionesDetalle(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                                        ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String, _
                                        Optional ByVal psCodCmac As String, Optional ByVal nTipoConcepto As Integer) As adodb.Recordset
    Dim sSql As String
    Dim rs As adodb.Recordset
    Dim lsFiltroUser As String
    Dim lsFiltroAge As String
    Dim lsFiltroTpoConcepto As String
    Dim oConect As COMConecta.DCOMConecta
    Dim lnTipoGrupo As GruposIngEgre
    
    
    Dim lsTablaDiaria As String
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    
    If pnTipoGrupo = gGruposIngEgreOtraAgencia Then
        lnTipoGrupo = gGruposIngEgreAgeLocal
    Else
        lnTipoGrupo = pnTipoGrupo
    End If
    
    If pdFecha = pdFechaDia Then
        lsTablaDiaria = " MovDiario "
    Else
        Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
        lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
    End If
    
    Set rs = New adodb.Recordset
    If psCodUser <> "" Then
        lsFiltroUser = " AND RIGHT(M.cMovNro,4)  = '" & psCodUser & "' "
    End If
    
    If nTipoConcepto = 1 Then
        lsFiltroTpoConcepto = "WHERE cGrupoNombre  like '%Aho.%'"
    End If

    sSql = " SELECT * FROM (" _
            & "SELECT CASE WHEN OpeCol.cCodAge = '" & psCodAge & "' THEN '000' ELSE (OpeCol.cCodAge + '2') END as nAgeOrd, OpeCol.cCodAge, cGrupoNombre, " _
            & "  (CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
            & "  (CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END ) AS Cheque , " _
            & "  (CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, G.cGrupoCod " _
            & "FROM GrupoOpe G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
            & "JOIN (SELECT SUBSTRING(MC.cCtaCod,4,2) AS cCodAge, Mc.cOpeCod, O.cOpeDesc, " _
            & " CASE WHEN (MD.nDocTpo IS NULL OR MC.COPECOD LIKE '99%') THEN (MC.nMonto) ELSE 0 END AS Efectivo , " _
            & " CASE WHEN (MD.nDocTpo = " & TpoDocCheque & " And MC.cOpeCod NOT LIKE '99%') THEN  (MC.nMonto) ELSE 0 END AS Cheque, " _
            & " CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " AND NOT MC.COPECOD LIKE '99%' THEN (MC.nMonto) ELSE 0 END AS OrdenPago " _
            & " FROM " & lsTablaDiaria & " M JOIN MovCap MC ON MC.nMovNro = M.nMovNro JOIN " _
            & " (Select nMovNro, cOpeCod, cCtaCod From MovCapDet Where nConceptoCod NOT IN (21,23) ) MCD ON MC.nMovNro = MCD.nMovNro And MC.cCtaCod = MCD.cCtaCod And MC.cOpeCod = MCD.cOpeCod " _
            & " LEFT JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
            & " JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
            & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
            & " AND SUBSTRING(MC.cCtaCod,9,1)='" & pnMoneda & "' And M.nMovFlag =" & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' "
            
         sSql = sSql & "UNION ALL " _
            & " SELECT SUBSTRING(M.cMovNro,18,2) AS cCodAge, M.cOpeCod, O.cOpeDesc," _
            & " CASE WHEN MD.nDocTpo IS NULL THEN (MV.nMovImporte) ELSE 0 END AS Efectivo , " _
            & " CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN (MV.nMovImporte) ELSE 0 END AS Cheque," _
            & " CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN (MV.nMovImporte) ELSE 0 END AS OrdenPago " _
            & " FROM " & lsTablaDiaria & " M " _
            & " JOIN MovOpevarias MV ON MV.nMovNro = M.nMovNro" _
            & " JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ")" _
            & " JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod" _
            & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "'" & lsFiltroUser & " " _
            & " AND MV.nMoneda = '" & pnMoneda & "' And M.nMovFlag = 0 AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' "
            
        sSql = sSql & ") As OpeCol  " _
            & " ON OpeCol.cOpeCod = GO.cOpeCod " _
            & " WHERE G.cGrupoCod LIKE '" & Format(lnTipoGrupo, "00") & "%' "
        sSql = sSql + " UNION ALL " _
            & " (SELECT DISTINCT CASE WHEN cAgeCod = '" & psCodAge & "' THEN '000' ELSE  (cAgeCod + '1') END as nAgeOrd, cAgeCod, cAgeDescripcion,0,0,0,'' " _
            & " FROM Agencias AG JOIN " _
            & " (SELECT SUBSTRING(MC.cCtaCod,4,2) As cCodAge " _
            & "  FROM MovCap MC JOIN " & lsTablaDiaria & " M ON M.nMovNro = MC.nMovNro " _
            & "   JOIN (SELECT NMOVNRO,COPECOD,CCTACOD,NCONCEPTOCOD FROM MOVCAPDET WHERE NCONCEPTOCOD<>21) MCD ON   MCD.NMOVNRO=MC.NMOVNRO AND MCD.CCTACOD=MC.CCTACOD AND MCD.COPECOD=MC.COPECOD " _
            & "  WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
            & "  AND SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
            & " ) Mov ON Mov.cCodAge = Ag.cAgeCod " & IIf(pnTipoGrupo = gGruposIngEgreAgeLocal, " And cAgeCod = '" & psCodAge & "'", " And cAgeCod <> '" & psCodAge & "' ") & " ) " _
            & " ) X " _
            & "Where " & IIf(pnTipoGrupo = gGruposIngEgreAgeLocal, " cCodAge = '" & psCodAge & "'", " cCodAge <> '" & psCodAge & "'")
    
   Set rs = oConect.CargaRecordSet(sSql)
    Set GetOpeIngEgreCaptacionesDetalle = rs
    If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
    End If
    oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetOpeIngEgreColocacionesDetalle(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
            ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String, _
            Optional ByVal psCodCmac As String, Optional ByVal nTipoConcepto As Integer) As adodb.Recordset

    Dim sSql As String
    Dim rs As adodb.Recordset
    Dim lsFiltroUser As String
    Dim lsFiltroAge As String
    Dim lsFiltroTpoConcepto As String
    Dim oConect As COMConecta.DCOMConecta
    Dim lsTablaDiaria As String
    Dim lnTipoGrupo  As GruposIngEgre
    
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    
    
        If pnTipoGrupo = gGruposIngEgreOtraAgencia Then
            lnTipoGrupo = gGruposIngEgreAgeLocal
        Else
            lnTipoGrupo = pnTipoGrupo
        End If
    
    
    If pdFecha = pdFechaDia Then
        lsTablaDiaria = " MovDiario "
    Else
        'psCodUser = IIf(psCodUser = "", "9999", psCodUser)
        Call OpeDiaCreaTemporal(Format(pdFecha, "yyyymmdd"), pgsCodUser, oConect)
        lsTablaDiaria = " [dbo].[##Mov_" & pgsCodUser & "] "
    End If
    
    Set rs = New adodb.Recordset
    
    If psCodUser <> "" Then
        lsFiltroUser = " AND RIGHT(M.cMovNro,4)  = '" & psCodUser & "' "
    End If
    
    If nTipoConcepto = 1 Then
        lsFiltroTpoConcepto = "WHERE cGrupoCod like '0[0-5]3%' and (cGrupoCod not like '0030[45]001' and cGrupoCod not like '013[01][138]001' and cGrupoNombre not like 'Pig.%')"
    ElseIf nTipoConcepto = 2 Then
        lsFiltroTpoConcepto = "WHERE cGrupoNombre like 'Pig.%'"
    ElseIf nTipoConcepto = 3 Then
        lsFiltroTpoConcepto = "WHERE (cGrupoCod like '0030[45]001' or cGrupoCod like '013[01][138]001')"
    End If
    
    sSql = "SELECT * FROM (" _
        & "SELECT CASE WHEN OpeCol.cCodAge = '" & psCodAge & "' THEN '000' ELSE (OpeCol.cCodAge + '2') END as nAgeOrd, OpeCol.cCodAge, cGrupoNombre, " _
        & "  (CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
        & "  (CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END ) AS Cheque, " _
        & "  (CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, G.cGrupoCod " _
        & "FROM GrupoOpe G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
        & "JOIN (SELECT SUBSTRING(MC.cCtaCod,4,2) AS cCodAge, MC.cOpeCod, O.cOpeDesc, " _
        & "  CASE WHEN MD.nDocTpo IS NULL THEN (MC.nMonto) ELSE 0 END AS Efectivo, " _
        & "  CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN  (MC.nMonto) ELSE 0 END AS Cheque, " _
        & "  CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN (MC.nMonto) ELSE 0 END AS OrdenPago " _
        & "  FROM " & lsTablaDiaria & " M JOIN MovCol MC ON MC.nMovNro = M.nMovNro " _
        & "  LEFT JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
        & "  JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
        & "  WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
        & "  AND SUBSTRING(MC.cCtaCod,9,1)='" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
        & "  ) As OpeCol  " _
        & "ON OpeCol.cOpeCod = GO.cOpeCod " _
        & "WHERE G.cGrupoCod LIKE '" & Format(lnTipoGrupo, "00") & "%' "
    sSql = sSql + " UNION ALL " _
        & "(SELECT DISTINCT CASE WHEN cAgeCod = '" & psCodAge & "' THEN '000' ELSE  (cAgeCod + '1') END as nAgeOrd, cAgeCod, cAgeDescripcion,0,0,0,'' " _
        & "FROM Agencias AG JOIN " _
        & "(SELECT SUBSTRING(MC.cCtaCod,4,2) As cCodAge " _
        & " FROM MovCol MC JOIN " & lsTablaDiaria & " M ON M.nMovNro = MC.nMovNro " _
        & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser & lsFiltroAge _
        & " AND SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And M.nMovFlag =" & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
        & ") Mov ON Mov.cCodAge = Ag.cAgeCod   " & IIf(pnTipoGrupo = gGruposIngEgreAgeLocal, " And cAgeCod = '" & psCodAge & "'", " And cAgeCod <> '" & psCodAge & "' ") & ") " _
        & " ) X " _
        & lsFiltroTpoConcepto _
        & " AND " & IIf(pnTipoGrupo = gGruposIngEgreAgeLocal, " cCodAge = '" & psCodAge & "'", " cCodAge <> '" & psCodAge & "'")
    
    Set rs = oConect.CargaRecordSet(sSql)
    Set GetOpeIngEgreColocacionesDetalle = rs
    If pdFecha <> pdFechaDia Then
            Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
    End If
    oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetOpeIngEgreOpeCMACsDetalle(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, ByVal psCodAge As String, _
            ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String, Optional ByVal nTipoConcepto As Integer) As adodb.Recordset

Dim sSql As String
Dim rs As adodb.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim lsFiltroTpoConcepto As String
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
    
    
Set rs = New adodb.Recordset
If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.cMovNro,4)  = '" & psCodUser & "' "
End If

If nTipoConcepto = 1 Then
    lsFiltroTpoConcepto = "WHERE cGrupoCod like '0[0-5]3%' and (cGrupoCod not like '0030[45]001' or cGrupoCod not like '013[01][138]001')"
ElseIf nTipoConcepto = 2 Then
    lsFiltroTpoConcepto = "WHERE cGrupoNombre like 'Pig.%'"
ElseIf nTipoConcepto = 3 Then
    lsFiltroTpoConcepto = "WHERE cGrupoCod like '0030[45]001' or cGrupoCod like '013[01][138]001'"
ElseIf nTipoConcepto = 1 Then
    lsFiltroTpoConcepto = "WHERE cGrupoNombre  like '%Aho.%'"
End If

sSql = "SELECT * FROM (" _
    & "SELECT (OpeCol.cCodAge + '2') As nAgeOrd, OpeCol.cCodAge, cGrupoNombre, " _
    & "  (CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
    & "  (CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END ) AS Cheque, " _
    & "  (CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, G.cGrupoCod " _
    & "FROM GrupoOpe G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
    & "JOIN (SELECT MC.cPersCod As cCodAge, M.cOpeCod, O.cOpeDesc, " _
    & "  CASE WHEN MD.nDocTpo IS NULL THEN (MC.nMonto) ELSE 0 END AS Efectivo, " _
    & "  CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN  (MC.nMonto) ELSE 0 END AS Cheque, " _
    & "  CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN (MC.nMonto) ELSE 0 END AS OrdenPago " _
    & "  FROM " & lsTablaDiaria & " M JOIN MovCMAC MC ON MC.nMovNro = M.nMovNro " _
    & "  LEFT JOIN MovDoc MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
    & "  JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod " _
    & "  WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
    & "  AND MC.nMoneda = '" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
    & "  ) As OpeCol  " _
    & "ON OpeCol.cOpeCod = GO.cOpeCod " _
    & "WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%' And G.nEfectivo = 1 "
sSql = sSql + " UNION " _
    & "(SELECT DISTINCT (P.cPersCod + '1') As nAgeOrd, P.cPersCod As cAgeCod, P.cPersNombre As cAgeDescripcion,0,0,0,'' " _
    & "FROM Persona P JOIN " _
    & "(SELECT MC.cPersCod As cCodAge " _
    & " FROM MovCMAC MC JOIN " & lsTablaDiaria & " M ON M.nMovNro = MC.nMovNro " _
    & " WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser _
    & " AND MC.nMoneda = '" & pnMoneda & "' And M.nMovFlag = " & gMovFlagVigente & " AND SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' " _
    & ") Mov ON Mov.cCodAge = P.cPersCod) " _
    & " ) X " _
    & lsFiltroTpoConcepto

Set rs = oConect.CargaRecordSet(sSql)
Set GetOpeIngEgreOpeCMACsDetalle = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetOpeIngEgreOtrasOpeDetalle(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String, _
                Optional ByVal nTipoConcepto As Integer) As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim lsFiltroUser2 As String
Dim lsFiltroAge2 As String
Dim lsFiltroTpoConcepto As String
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

Set rs = New adodb.Recordset
If psCodAge <> "" Then
    lsFiltroAge = " AND  SUBSTRING(M.CMOVNRO,18,2) = '" & psCodAge & "'"
    lsFiltroAge2 = " AND  SUBSTRING(CMOVNRO,18,2) = '" & psCodAge & "'"
End If
If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.CMOVNRO,4) = '" & psCodUser & "' "
    lsFiltroUser2 = " AND RIGHT(CMOVNRO,4) = '" & psCodUser & "' "
End If

If nTipoConcepto = 1 Then
    lsFiltroTpoConcepto = "WHERE cGrupoCod like '0[0-5]3%' and (cGrupoCod not like '0030[45]001' or cGrupoCod not like '013[01][138]001')"
ElseIf nTipoConcepto = 2 Then
    lsFiltroTpoConcepto = "WHERE cGrupoNombre like 'Pig.%'"
ElseIf nTipoConcepto = 3 Then
    lsFiltroTpoConcepto = "WHERE cGrupoCod like '0030[45]001' or cGrupoCod like '013[01][138]001'"
ElseIf nTipoConcepto = 1 Then
    lsFiltroTpoConcepto = "WHERE cGrupoNombre  like '%Aho.%'"
End If

'--Otras Operaciones
sql = "SELECT * FROM (" _
    & "SELECT cGrupoNombre, cCodAge = '" & psCodAge & "', " _
    & "  (CASE WHEN G.cIngEgr = 'I' THEN EFECTIVO ELSE EFECTIVO*-1 END) AS EFECTIVO, " _
    & "  (CASE WHEN G.cIngEgr = 'I' THEN CHEQUE ELSE CHEQUE*-1 END) AS CHEQUE, " _
    & "  (CASE WHEN G.cIngEgr = 'I' THEN ORDENPAGO ELSE ORDENPAGO*-1 END) AS ORDENPAGO, GO.cGrupoCod " _
    & "FROM GRUPOOPE  AS G JOIN GRUPOSOPE GO ON GO.cGrupoCod=G.cGrupoCod " _
    & "INNER JOIN (SELECT M.cOpeCod, O.COPEDESC, " _
    & "  CASE WHEN MD.nDocTpo IS NULL THEN  (MV.nMovImporte) ELSE 0 END AS EFECTIVO, " _
    & "  CASE WHEN MD.nDocTpo = " & TpoDocCheque & " THEN  (MV.nMovImporte) ELSE 0 END AS CHEQUE, " _
    & "  CASE WHEN MD.nDocTpo = " & TpoDocOrdenPago & " THEN  (MV.nMovImporte) ELSE 0 END AS ORDENPAGO " _
    & "  FROM " & lsTablaDiaria & " M JOIN MOVOPEVARIAS MV ON MV.nMovNro = M.nMovNro " _
    & "  LEFT JOIN MOVDOC MD ON MD.nMovNro = M.nMovNro AND MD.nDocTpo IN (" & TpoDocCheque & "," & TpoDocOrdenPago & ") " _
    & "  JOIN OPETPO O ON O.cOpeCod = M.cOpeCod " _
    & "  WHERE   SUBSTRING(M.CMOVNRO,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser _
    & "  AND M.NMOVFLAG =" & gMovFlagVigente & " AND MV.nMoneda =" & pnMoneda & " " _
    & "  ) AS OPESERV " _
    & "  ON OPESERV.cOpeCod = GO.cOpeCod " _
    & " WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%'  " _
    & " ) X " _
    & lsFiltroTpoConcepto

Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreOtrasOpeDetalle = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetOpeIngEgreServiciosDetalle(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, _
                pdFechaDia As Date, pgsCodUser As String, Optional ByVal nTipoConcepto As Integer) As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim lsFiltroTpoConcepto As String
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

Set rs = New adodb.Recordset
If psCodAge <> "" Then
    lsFiltroAge = " AND  SUBSTRING(M.CMOVNRO,18,2) = '" & psCodAge & "'"
End If
If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.CMOVNRO,4) = '" & psCodUser & "' "
End If

If nTipoConcepto = 1 Then
    lsFiltroTpoConcepto = "WHERE cGrupoCod like '0400[56]001%'"
ElseIf nTipoConcepto = 2 Then
    lsFiltroTpoConcepto = "WHERE cGrupoCod like '04%' and cGrupoCod not like '0400[56]001%'"
End If

'--Servicios
sql = "SELECT * FROM (" _
    & "SELECT cGrupoNombre, cCodAge = '" & psCodAge & "', " _
    & "  (CASE WHEN G.cIngEgr = 'I' THEN Efectivo ELSE Efectivo*-1 END) AS Efectivo, " _
    & "  (CASE WHEN G.cIngEgr = 'I' THEN Cheque ELSE Cheque*-1 END) AS Cheque, " _
    & "  (CASE WHEN G.cIngEgr = 'I' THEN OrdenPago ELSE OrdenPago*-1 END) AS OrdenPago, GO.cGrupoCod " _
    & "FROM GRUPOOPE AS G JOIN GRUPOSOPE GO ON GO.cGrupoCod = G.cGrupoCod " _
    & "INNER JOIN (SELECT M.cOpeCod, O.COPEDESC, (MV.nMonto) Efectivo, " _
    & "  Cheque = 0, OrdenPago = 0 FROM " & lsTablaDiaria & " M JOIN MovServicios MV ON MV.nMovNro = M.nMovNro and MV.cOpecod <> '" & gITFGiroCancelEfect & "'" _
    & "  JOIN OPETPO O ON O.cOpeCod = M.cOpeCod " _
    & "  WHERE SUBSTRING(M.CMOVNRO,1,8) = '" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser _
    & "  AND M.NMOVFLAG = " & gMovFlagVigente & " And MV.nMoneda = " & pnMoneda & " " _
    & "  ) AS OPESERV " _
    & "  ON OPESERV.cOpeCod = GO.cOpeCod " _
    & " WHERE G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%'  " _
    & " ) X " _
    & lsFiltroTpoConcepto
    'NAGL 20190307 Agregó and MV.cOpecod <> '" & gITFGiroCancelEfect & "'"
Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreServiciosDetalle = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetOpeIngEgreCompraVentaDetalle(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                        ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, _
                        pdFechaDia As Date, pgsCodUser As String, Optional ByVal nTipoConcepto As Integer) As adodb.Recordset

Dim sql As String
Dim rs As adodb.Recordset
Dim lsFiltroUser As String
Dim lsFiltroAge As String
Dim lsFiltroTpoConcepto As String
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

Set rs = New adodb.Recordset
If psCodAge <> "" Then
    lsFiltroAge = " AND  SUBSTRING(M.CMOVNRO,18,2) ='" & psCodAge & "'"
End If
If psCodUser <> "" Then
    lsFiltroUser = " AND RIGHT(M.CMOVNRO,4) = '" & psCodUser & "' "
End If

sql = "SELECT cGrupoNombre, cCodAge = '" & psCodAge & "', " _
    & IIf(pnMoneda = gMonedaNacional, " (EfecSoles) ", " (EfecDol) ") & " AS Efectivo, (Cheque) Cheque, " _
    & "(OrdenPago) OrdenPago, GO.cGrupoCod " _
    & "FROM GrupoOpe AS G JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod " _
    & "INNER JOIN (SELECT cOpeCod, cOpeDesc, EfecSoles, EfecDol, 0 AS CHEQUE, 0 AS ORDENPAGO " _
    & "  FROM (SELECT cOpeCod, cOpeDesc, " _
    & "  CASE WHEN cOpeCod IN ('" & gOpeCajeroMECompra & "','" & gOpeCajeroMECompraEsp & "') THEN EfectivoSoles*-1 ELSE EfectivoSoles END AS EfecSoles, " _
    & "  CASE WHEN cOpeCod IN ('" & gOpeCajeroMECompra & "','" & gOpeCajeroMECompraEsp & "') THEN EfectivoDolares ELSE EfectivoDolares*-1 END AS EfecDol " _
    & "  From (SELECT M.cOpeCod, O.cOpeDesc, " _
    & "        (ROUND(MC.nMovImporte*TC.nMovTpoCambio,2)) AS  EfectivoSoles, " _
    & "        (MC.nMovImporte) As EfectivoDolares " _
    & "        FROM " & lsTablaDiaria & " M JOIN MOVCOMPRAVENTA MC ON MC.nMovNro = M.nMovNro " _
    & "        JOIN MOVTPOCAMBIO TC ON TC.nMovNro = M.nMovNro " _
    & "        JOIN OPETPO O ON O.cOpeCod = M.cOpeCod " _
    & "        WHERE SUBSTRING(M.CMOVNRO,1,8) ='" & Format$(pdFecha, "yyyymmdd") & "' " & lsFiltroAge & lsFiltroUser _
    & "        AND M.NMOVFLAG = " & gMovFlagVigente & " " _
    & "     ) AS MTC ) AS TC) AS MTC " _
    & " ON MTC.cOpeCod = GO.cOpeCod"

Set rs = oConect.CargaRecordSet(sql)
Set GetOpeIngEgreCompraVentaDetalle = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetOpeIngEgreHabDevCajeroDetalle(ByVal pnTipoGrupo As GruposIngEgre, ByVal pdFecha As Date, _
                ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMoneda As Moneda, pdFechaDia As Date, pgsCodUser As String) As adodb.Recordset

Dim sSql As String
Dim rs As adodb.Recordset
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

Set rs = New adodb.Recordset


sSql = "SELECT cGrupoNombre, GO.cGrupoCod, GO.cOpeCod, cCodAge = '" & psCodAge & "', " _
    & "Efectivo = CASE WHEN G.cIngEgr = 'I' THEN (nMontoHabDev) ELSE (nMontoHabDev)*-1 END, " _
    & "(Cheque) As Cheque, (OrdenPago) As OrdenPago FROM GrupoOpe G JOIN GruposOpe GO " _
    & "ON GO.cGrupoCod = G.cGrupoCod JOIN " _
    & "(SELECT cOpeCod, Efectivo AS nMontoHabDev, 0 AS Cheque , 0 AS OrdenPago FROM " _
    & "(SELECT M1.cOpeCod, (MH.nMovImporte) AS Efectivo " _
    & "FROM OpeTpo O JOIN " & lsTablaDiaria & " M JOIN MovHabDevCajero MH ON MH.nMovNro = M.nMovNro ON O.cOpeCod = M.cOpeCod " _
    & "JOIN MovRef MR JOIN  " & lsTablaDiaria & "  M1 ON MR.nMovNro = M1.nMovNro ON MR.nMovNroRef = M.nMovNro " _
    & "WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format$(pdFecha, "yyyymmdd") & "' AND MH.cUsuDest = '" & psCodUser & "' AND " _
    & "M.nMovFlag = 0 AND MH.cAgeCod = '" & psCodAge & "' And MH.nMoneda = " & pnMoneda & " And M1.nMovFlag = " & gMovFlagVigente & " " _
    & "UNION " _
    & "SELECT M.cOpeCod, (MH.nMovImporte) AS Efectivo " _
    & "FROM OpeTpo O JOIN " & lsTablaDiaria & " M JOIN MovHabDevCajero MH ON MH.nMovNro = M.nMovNro ON O.cOpeCod = M.cOpeCod " _
    & "JOIN MovRef MR JOIN " & lsTablaDiaria & " M1 ON MR.nMovNro = M1.nMovNro ON MR.nMovNroRef = M.nMovNro " _
    & "WHERE SUBSTRING(M.cMovNro,1,8) = '" & Format$(pdFecha, "yyyymmdd") & "' AND MH.cUsuOrig = '" & psCodUser & "' AND " _
    & "M.nMovFlag = 0 AND MH.cAgeCod = '" & psCodAge & "' And MH.nMoneda = " & pnMoneda & " And M1.nMovFlag = " & gMovFlagVigente & " " _
    & ") HabDev) OpeHD ON OpeHD.cOpeCod = GO.cOpeCod Where " _
    & "G.cGrupoCod LIKE '" & Format(pnTipoGrupo, "00") & "%' "

Set rs = oConect.CargaRecordSet(sSql)
Set GetOpeIngEgreHabDevCajeroDetalle = rs
If pdFecha <> pdFechaDia Then
        Call OpeDiaEliminaTemporal(pgsCodUser, oConect)
End If

oConect.CierraConexion: Set oConect = Nothing
End Function
'** End Juez *************************************************************************
'MIOL 20130308, SEGUN ERS008 - 2013 **************************************************
Public Function GetBuscarConfirmacionEfec(ByVal psCodUser As String, ByVal psFecha As String, ByVal psCodAge As String) As Boolean
Dim sSql As String, rs As adodb.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New adodb.Recordset
If oCon.AbreConexion = False Then Exit Function

GetBuscarConfirmacionEfec = False

    sSql = "SELECT *    FROM OpeTpo O "
    sSql = sSql & "     JOIN Mov M On M.cOpeCod=O.cOpeCod "
    sSql = sSql & "     JOIN MovHabDevCajero H On H.nMovNro=M.nMovNro "
    sSql = sSql & "     JOIN RRHH RH On H.cUsuOrig = RH.cUser "
    sSql = sSql & "     LEFT JOIN ( Select MR.nMovNroRef "
    sSql = sSql & "                 From MovRef MR "
    sSql = sSql & "                         JOIN Mov M1 ON MR.nMovNro = M1.nMovNro "
    sSql = sSql & "                 Where M1.nMovFlag NOT IN (2,3,1,5) "
    sSql = sSql & "               ) MR ON M.nMovNro = MR.nMovNroRef "
    sSql = sSql & "     WHERE M.cOpeCod IN ('901013') AND M.nMovFlag NOT IN (2,3,1,5) AND H.cAgeCod = '" & psCodAge & "' "
    sSql = sSql & "  AND M.cMovNro Like '" & psFecha & "%' AND H.cUsuDest = '" & psCodUser & "'  And MR.nMovNroRef IS NULL "
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF Then
        GetBuscarConfirmacionEfec = True
    End If
    
    Set rs = Nothing
    oCon.CierraConexion
    Set oCon = Nothing

End Function
'END MIOL ****************************************************************************


' ****************************************** Agregado por RIRO el 20130403 - Proyecto de recaudos

Public Function GetOpeIngEgreRecaudo(ByVal pdFecha As Date, _
                                     ByVal psCodAge As String, _
                                     ByVal psCodUser As String, _
                                     ByVal pnMoneda As Moneda, _
                                     ByVal pdFechaDia As Date) As adodb.Recordset
    Dim sSql As String
    Dim rs As adodb.Recordset
    
    Dim oConect As COMConecta.DCOMConecta
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    
    Set rs = New adodb.Recordset
    
    sSql = "stp_sel_IngrEgrRecaudo '" & psCodUser & "', '"
    sSql = sSql & Format$(pdFecha, "yyyymmdd") & "' ,'" & psCodAge & "', " & pnMoneda
    
    Set rs = oConect.CargaRecordSet(sSql)
    Set GetOpeIngEgreRecaudo = rs
    
    oConect.CierraConexion: Set oConect = Nothing

End Function

' Fin RIRO *********************************************************************************

'*** RIRO 20130702 SEGUN TI-ERS083-2013 *******

Public Function ResumenEgresosEfectivo(ByVal sFecha As String, ByVal sUser As String, ByVal sAgencia As String) As adodb.Recordset
    
    Dim sSql As String
    Dim rs As adodb.Recordset
    Dim oConect As COMConecta.DCOMConecta
    
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    Set rs = New adodb.Recordset
        
    sSql = "stp_sel_ResumenEgresosEfectivo '" & sFecha & "', '" & sUser & "', '" & sAgencia & "'"
        
    Set rs = oConect.CargaRecordSet(sSql)
    Set ResumenEgresosEfectivo = rs
    
    oConect.CierraConexion: Set oConect = Nothing

End Function

Public Function ReporteEgresos(ByVal sAgencia As String, ByVal sDesde As String, ByVal sHasta As String) As adodb.Recordset

    Dim sSql As String
    Dim rs As adodb.Recordset
    Dim oConect As COMConecta.DCOMConecta
    
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    Set rs = New adodb.Recordset
        
    sSql = "stp_sel_ReporteEgresos '" & sAgencia & "', '" & sDesde & "', '" & sHasta & "'"
        
    Set rs = oConect.CargaRecordSet(sSql)
    
    Set ReporteEgresos = rs
    
    oConect.CierraConexion: Set oConect = Nothing
    
End Function

'*** Fin RIRO *********************************
'FRHU 20141203 ERS048-2014
Public Function GetOpeIngEgreNoEfectivo(ByVal pdFecha As Date, ByVal psCodAge As String, ByVal psCodUser As String, _
                                     ByVal pnMoneda As Moneda, ByVal pdFechaDia As Date) As adodb.Recordset
    Dim sSql As String
    Dim rs As adodb.Recordset
    Dim oConect As COMConecta.DCOMConecta
    
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    Set rs = New adodb.Recordset
    
    sSql = "stp_sel_IngEgeNoEfectivo '" & psCodUser & "', '" & Format$(pdFecha, "yyyymmdd") & "' ,'" & psCodAge & "', " & pnMoneda
    
    Set rs = oConect.CargaRecordSet(sSql)
    Set GetOpeIngEgreNoEfectivo = rs
    oConect.CierraConexion
    Set oConect = Nothing
End Function
'FIN FRHU 20141203
'PASI
Public Function ObtieneUserxRegularizaFaltante(ByVal pnMovNro As Long) As String
    Dim sSql As String
    Dim rs As adodb.Recordset
    Dim oConect As COMConecta.DCOMConecta
    
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    Set rs = New adodb.Recordset
    
    sSql = "stp_sel_ObtieneUserxRegularizaFaltante " & pnMovNro
    
    Set rs = oConect.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
        ObtieneUserxRegularizaFaltante = rs!cUser
    Else
        ObtieneUserxRegularizaFaltante = ""
    End If
    oConect.CierraConexion
    Set oConect = Nothing
End Function
'end pasi
'FRHU 20160220
Public Function ObtieneRegistroBilletaje(ByVal pnMovNroRegEfec As Long) As adodb.Recordset
    Dim sSql As String
    Dim rs As adodb.Recordset
    Dim oConect As COMConecta.DCOMConecta
    
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    Set rs = New adodb.Recordset
    
    sSql = "stp_sel_ObtenerRegistroBilletaje " & pnMovNroRegEfec
    
    Set rs = oConect.CargaRecordSet(sSql)
    Set ObtieneRegistroBilletaje = rs
    oConect.CierraConexion
    Set oConect = Nothing
End Function
'FIN FRHU
'**GIPO ERS051-2016
Public Function ObtenerTarjetasADevolver(ByVal user As String, ByVal fecha As Date, psCodAge As String) As adodb.Recordset
Dim loConecta As COMConecta.DCOMConecta
Dim lsSql As String
    
    Set loConecta = New COMConecta.DCOMConecta
    If loConecta.AbreConexion = False Then Exit Function

    lsSql = "exec DBTarjeta..stp_sel_VISARecuperaTarjetasDevolver '" & user & "','" & Format(fecha, "yyyymmdd") & "','" & psCodAge & "'"
 
    Set ObtenerTarjetasADevolver = loConecta.CargaRecordSet(lsSql)
    loConecta.CierraConexion
    Set loConecta = Nothing
End Function







