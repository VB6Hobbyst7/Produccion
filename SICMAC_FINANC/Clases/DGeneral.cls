VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long

Private dbCmact As DConecta
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String

Public Function GetCOM(ByVal sDispositivo As String) As String
Dim rsPC As Recordset
Dim sSql As String
Dim sMaquina As String

GetCOM = ""
sMaquina = GetNombreComputadora
sSql = "Select cPuerto from SetupCOM where cDispositivo = '" & sDispositivo & "' " _
    & "and cNomPC = '" & sMaquina & "'"

Set rsPC = dbCmact.CargaRecordSet(sSql)
If Not (rsPC.EOF And rsPC.BOF) Then
    GetCOM = "COM" & Trim(rsPC!cPuerto)
End If
RSClose rsPC
End Function

Public Function GetNombreComputadora() As String
Dim sbuffMaq As String
Dim nSizeMaq As Long
sbuffMaq = Space(255)
nSizeMaq = Len(sbuffMaq)
GetComputerName sbuffMaq, nSizeMaq
GetNombreComputadora = Trim(Left$(sbuffMaq, nSizeMaq))
End Function

Public Function GetPersonaEdad(ByVal sPersona As String) As Long
Dim rsPers As Recordset
Dim sSql As String
sSql = "SELECT DATEDIFF(yy,dPersNacCreac,Getdate()) Edad FROM " & sDBPersona & "Persona " _
    & "WHERE cPersCod = '" & sPersona & "'"
Set rsPers = dbCmact.CargaRecordSet(sSql)
If rsPers.EOF And rsPers.BOF Then
    GetPersonaEdad = 0
Else
    GetPersonaEdad = rsPers("Edad")
End If
rsPers.Close
Set rsPers = Nothing
End Function

Public Function GetOperacionesUsuario(ByVal sUsuario As String, ByVal sProducto As String, ByVal MatOpe As Variant, ByVal NroRegOpe As Integer, Optional ByVal psMoneda As String = "") As Recordset
Dim rsVar As Recordset
Dim sSql As String
Dim lsFiltroMon As String
Dim i As Integer
Dim nCont As Integer
Dim sCodOpe As String
If psMoneda <> "" Then
    lsFiltroMon = " AND O.cOpeCod NOT like '__" & psMoneda & "%'"
End If
sSql = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O " _
    & "WHERE O.cOpeCod LIKE '" & sProducto & "%' And O.cOpeVisible = '1'  " & lsFiltroMon _
    & " Order by O.cOpeCod, O.nOpeNiv "
Set rsVar = dbCmact.CargaRecordSet(sSql)
sCodOpe = ""
nCont = -1
Do While Not rsVar.EOF
    For i = nCont + 1 To NroRegOpe - 1
        If Trim(rsVar!cOpeCod) = Trim(MatOpe(i, 0)) Then
            sCodOpe = sCodOpe & Trim(rsVar!cOpeCod) & "','"
            nCont = i
            Exit For
        End If
    Next i
    rsVar.MoveNext
Loop
    If sCodOpe <> "" Then
        sCodOpe = Mid(sCodOpe, 1, Len(sCodOpe) - 2)
    End If

sSql = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O "
If sCodOpe <> "" Then
    sSql = sSql & " Where O.cOpeCod in ('" & sCodOpe & ")"
Else
    sSql = sSql & " Where O.cOpeCod in ('')"
End If
Set rsVar = dbCmact.CargaRecordSet(sSql)
Set GetOperacionesUsuario = rsVar
Set rsVar = Nothing
End Function

Public Function GetOperacionesUsuario1(ByVal sUsuario As String, ByVal sProducto As String, ByVal MatOpe As Variant, ByVal NroRegOpe As Integer, ByVal psMoneda As String) As Recordset
Dim rsVar As Recordset
Dim sSql As String
Dim lsFiltroMon As String
Dim i As Integer
Dim nCont As Integer
Dim sCodOpe As String
If psMoneda <> "" Then
    lsFiltroMon = " AND O.cOpeCod NOT like '__" & psMoneda & "%'"
End If
sSql = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O " _
    & "WHERE O.cOpeCod LIKE '" & sProducto & "%' And O.cOpeVisible = '1'  " & lsFiltroMon _
    & " Order by O.cOpeCod, O.nOpeNiv "
Set rsVar = dbCmact.CargaRecordSet(sSql)
sCodOpe = ""
nCont = -1
Do While Not rsVar.EOF
    For i = nCont + 1 To NroRegOpe - 1
        If Trim(rsVar!cOpeCod) = Trim(MatOpe(i, 0)) Then
            sCodOpe = sCodOpe & Trim(rsVar!cOpeCod) & "','"
            nCont = i
            Exit For
        End If
    Next i
    rsVar.MoveNext
Loop
    If sCodOpe <> "" Then
        sCodOpe = Mid(sCodOpe, 1, Len(sCodOpe) - 2)
    End If

sSql = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O "
If sCodOpe <> "" Then
    sSql = sSql & " Where O.cOpeCod in ('" & sCodOpe & ")"
Else
    sSql = sSql & " Where O.cOpeCod in ('')"
End If
Set rsVar = dbCmact.CargaRecordSet(sSql)
Set GetOperacionesUsuario1 = rsVar
Set rsVar = Nothing
End Function

Public Function GetTrabajadorCMACT() As Recordset
Dim sSql As String
sSql = "Select P.cPersNombre, P.cPersCod FROM " & sDBPersona & "Persona P INNER JOIN RRHH T ON " _
    & "P.cPersCod = T.cPersCod Where cRHEstado NOT IN (" & gRHEstDespedido & "," _
    & gRHEstRetirado & ") ORDER BY cPersNombre"
Set GetTrabajadorCMACT = dbCmact.CargaRecordSet(sSql)
End Function

Public Function GeneraNuevaCuenta(ByVal sAgencia As String, ByVal nProducto As Producto, ByVal nMoneda As Moneda) As String
Dim rsCta As Recordset
Dim sCuenta As String, sCorr As String, sSql As String, sDigOK As String
sSql = "Select ISNULL(CONVERT(INT,MAX(SUBSTRING(cCtaCod,10,9))),0) + 1 Correlativo FROM " _
    & "Producto Where cCtaCod LIKE  '" & Trim(sAgencia) & Trim(nProducto) & Trim(nMoneda) & "%'"
Set rsCta = dbCmact.CargaRecordSet(sSql)
If Not (rsCta.EOF And rsCta.BOF) Then
    sCorr = FillNum(Trim(rsCta("Correlativo")), 9, "0")
    sDigOK = GetDigitoChequeo(sCorr)
    sCuenta = sAgencia & Trim(nMoneda) & sCorr & sDigOK
End If
RSClose rsCta
GeneraNuevaCuenta = sCuenta
End Function

Public Function GetDigitoChequeo(ByVal psCadena As String) As Integer
Dim liFactor As Variant
Dim liCadena(1 To 5) As Integer
Dim liSum, i As Integer
Dim lnDigito As Integer
liFactor = Array(6, 5, 4, 3, 2)
liCadena(1) = Val(Mid(psCadena, 1, 1))
liCadena(2) = Val(Mid(psCadena, 2, 1))
liCadena(3) = Val(Mid(psCadena, 3, 1))
liCadena(4) = Val(Mid(psCadena, 4, 1))
liCadena(5) = Val(Mid(psCadena, 5, 1))
liSum = 0
For i = 1 To 5
    liSum = liSum + liCadena(i) * liFactor(i - 1)
Next i
lnDigito = 11 - (liSum Mod 11)
If lnDigito = 10 Then
    GetDigitoChequeo = 0
ElseIf lnDigito = 11 Then
    GetDigitoChequeo = 1
Else
    GetDigitoChequeo = lnDigito
End If
End Function

Public Function GetConstante(ByVal nConstante As ConstanteCabecera) As Recordset
Dim sSql As String
Dim rsRel As Recordset
sSql = "SELECT UPPER(cConsDescripcion) cDescripcion, nConsValor FROM " & sDBComunes & "Constante WHERE " _
    & "nCONSCOD = " & nConstante & " AND nCONSCOD <> nConsValor ORDER BY nConsValor"
Set GetConstante = dbCmact.CargaRecordSet(sSql)
End Function
Public Function GetConstanteNiv(ByVal nConstante As ConstanteCabecera) As Recordset
Dim sSql As String
sSql = "SELECT nConsValor , UPPER(cConsDescripcion) cDescripcion, 1 as Nivel FROM " & sDBComunes & "Constante WHERE " _
    & "nCONSCOD = " & nConstante & " AND nCONSCOD <> nConsValor ORDER BY nConsValor"
Set GetConstanteNiv = dbCmact.CargaRecordSet(sSql)
End Function

Public Function GetDataUser(ByVal psCodDato As String, Optional pbBuscaXCodPers As Boolean = False) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Set oConect = New DConecta
Set rs = New ADODB.Recordset
If oConect.AbreConexion = False Then Exit Function

'WIOR 20130828 COMENTO Y SE AGREGO EN UN PROCEDIMIENTO ALMACENADO
'sql = " Select  P.cPersNombre , P.cPersDireccDomicilio  , RH.cPersCod,  Rh.cAreaCod , A.cAreaDescripcion, " _
'    & "         cAgenciaActual, ISNULL(Ag.cAgeDescripcion,'') as cDescAgActual , RH.cAreaCodActual, A1.cAreaDescripcion as cDescAreaActual , " _
'    & "         A.cAreaEstruc,  cAgenciaAsig, ISNULL(AG1.cAgeDescripcion,'') as cDescAgAsig , " _
'    & "         ISNULL((SELECT cPersIDnro FROM PERSID WHERE cPersIDTpo ='" & gPersIdDNI & "' AND cPersCod = p.cPersCod ),'') AS DNI, " _
'    & "         ISNULL((SELECT cPersIDnro FROM PERSID WHERE cPersIDTpo ='" & gPersIdRUC & "' AND cPersCod = p.cPersCod ),'') AS RUC,   " _
'    & "         ISNULL(C.cRHCargoCod,'') AS cRHCargoCod,  ISNULL(C.cRHCargoDescripcion,'') AS cRHCargoDescripcion, " _
'    & "         ISNULL(C.cConsDescripcion,'') AS cConsDescripcion, ISNULL(C.cRHCargoCategoria,'') cRHCargoCategoria , ISNULL(A.nLogTramite,0) AS  cLogTramite " _
'    & " From    RRHH RH JOIN Persona P on P.cPersCod = RH.cPersCod " _
'    & "         LEFT JOIN (SELECT dRHCargoFecha as FechaCargo,C.cPersCod, C.cRHCargoCod, CT.cRHCargoDescripcion , CC.cConsDescripcion, CT.cRHCargoCategoria " _
'    & "                    FROM     RHCargos C " _
'    & "                             JOIN RHCargosTabla CT   ON CT.cRHCargoCod = C.cRHCargoCod " _
'    & "                             JOIN CONSTANTE CC   ON CC.nConsValor  = CT.cRHCargoCategoria " _
'    & "                     WHERE  CC.nCONSCOD LIKE '" & gRHCategoria & "' " _
'    & "                    and dRHCargoFecha = (SELECT MaX(dRHCargoFecha) FROM RHCargos WHERE cPersCod = c.cPersCod) ) as C " _
'    & "         ON C.cPersCod = RH.cPersCod " _
'    & "         LEFT JOIN Areas A ON A.cAreaCod = RH.cAreaCod " _
'    & "         LEFT JOIN Agencias AG ON  AG.cAgeCod = LTRIM(RTRIM(RH.cAgenciaActual)) " _
'    & "         LEFT JOIN Agencias AG1 on AG1.cAgecod= LTRIM(RTRIM(RH.cAgenciaAsig )) " _
'    & "         LEFT JOIN Areas A1 ON A1.cAreaCod = RH.cAreaCodActual " _
'    & " Where   " & IIf(pbBuscaXCodPers = True, " RH.cPersCod ='" & psCodDato & "' ", " cUser ='" & psCodDato & "' ")

sql = "EXEC stp_sel_GetDataUser '" & psCodDato & "'" 'WIOR 20130828
Set rs = oConect.CargaRecordSet(sql)
Set GetDataUser = rs
oConect.CierraConexion
Set oConect = Nothing
End Function

Public Function CuentaNombre(ByVal pCodCta As String) As String
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    Dim oConect As DConecta
    
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then Exit Function
    
    pCodCta = Trim(pCodCta)
    If Len(pCodCta) > 4 Then
       tmpSql = "SELECT RTRIM(ISNULL(c1.cCtaContDesc,'')) + ' ' + RTRIM(c.cCtaContDesc) as cCtaContDesc " _
              & "FROM   " & sDBComunes & "CtaCont c LEFT JOIN " & sDBComunes & "CtaCont c1 ON c1.cCtaContCod = substring(c.cCtaContCod,1," & Len(pCodCta) - 2 & ") " _
              & "WHERE  c.cCtaContCod = '" & pCodCta & "'"
    Else
       tmpSql = "SELECT cCtaContDesc FROM " & sDBComunes & "CtaCont WHERE cCtaContCod = '" & pCodCta & "'"
    End If
    Set tmpReg = oConect.CargaRecordSet(tmpSql)
    If (tmpReg.BOF Or tmpReg.EOF) Then
        CuentaNombre = ""
    Else
        CuentaNombre = Trim(tmpReg!cCtaContDesc)
    End If
    tmpReg.Close: Set tmpReg = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim sConn As String
Dim ClsIni As DClassIni
Set ClsIni = New DClassIni
sConn = ClsIni.CadenaConexion
sDBComunes = ClsIni.BaseComunes
sDBPersona = ClsIni.BasePersonas
sDBImagenes = ClsIni.BaseImagenes
Set ClsIni = Nothing
Set dbCmact = New DConecta
dbCmact.AbreConexion
End Sub

Private Sub Class_Terminate()
dbCmact.CierraConexion
End Sub
Public Function GetUserAreaAgencia(ByVal psAreaCod As String, Optional ByVal psAgeCod As String, Optional psUser As String = "", Optional ByVal pbExcluir As Boolean = True) As ADODB.Recordset
Dim sql As String
Dim lsFiltro As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Set oConect = New DConecta
Set rs = New ADODB.Recordset
If oConect.AbreConexion = False Then Exit Function

If psUser <> "" Then
    lsFiltro = " AND RH.cUser " & IIf(pbExcluir, " NOT ", "") & " IN ('" & psUser & "')"
End If

sql = "SELECT   UPPER(RH.cUser) AS cUser, REPLACE(REPLACE(Replace(P.cPersNombre,'/',' '),'-',''),'\' ,' DE ') AS cNomUser , 1 Nivel  " _
    & " FROM    RRHH RH JOIN PERSONA P ON P.CPERSCOD = RH.CPERSCOD " _
    & " WHERE   RH.cAreaCodActual ='" & psAreaCod & "' AND RH.cAgenciaActual='" & psAgeCod & "' " _
    & "         AND RH.cUser<>'' " & lsFiltro _
    & " Order BY  P.cPersNombre"

Set rs = oConect.CargaRecordSet(sql)
Set GetUserAreaAgencia = rs
oConect.CierraConexion
Set oConect = Nothing
End Function


Public Function GetnMovNro(ByVal psMovNro As String) As Long
Dim lnMovNro As Long
Dim oCon As DConecta
Dim cmd As New ADODB.Command
Set oCon = New DConecta

If oCon.AbreConexion = False Then Exit Function
cmd.CommandText = "sp_GetnMovNro"
cmd.CommandType = adCmdStoredProc
cmd.Name = "spGetMovNro"
Set cmd.ActiveConnection = oCon.ConexionActiva
cmd.CommandTimeOut = 720
cmd.Parameters.Refresh
cmd.Parameters("@pnMovNro").Direction = adParamOutput
oCon.ConexionActiva.spGetMovNro Trim(psMovNro)
lnMovNro = cmd.Parameters("@pnMovNro")
Set cmd = Nothing
GetnMovNro = lnMovNro
oCon.CierraConexion
Set oCon = Nothing

End Function
Public Function GetParametro(ByVal pnProdParam As Long, ByVal nParametro As Long) As Double
Dim rsVar As Recordset
Dim sSql As String
Dim oConect As DConecta
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

sSql = "SELECT nParValor FROM Parametro WHERE nParCod = " & nParametro & " " _
    & "And nParProd = " & pnProdParam

GetParametro = 0
Set rsVar = New Recordset
Set rsVar = oConect.CargaRecordSet(sSql)
If Not rsVar.EOF And Not rsVar.BOF Then
    GetParametro = rsVar("nParValor")
End If
rsVar.Close
Set rsVar = Nothing
oConect.CierraConexion
Set oConect = Nothing
End Function
Public Function GetMensajeBoletas(ByVal psProducto As String) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
GetMensajeBoletas = ""
If oCon.AbreConexion = False Then Exit Function
sql = "Select cMensaje From Mensajes where cCodPro ='" & psProducto & "'"
Set rs = oCon.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetMensajeBoletas = rs!cMensaje
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion
Set oCon = Nothing
End Function


Public Function EmiteTipoCambio(ByVal dFecha As Date, ByVal nTpoTipoCambio As TipoCambio) As Double
    Dim sql As String
    Dim rs As Recordset
    EmiteTipoCambio = 0
    Set rs = New Recordset
    rs.CursorLocation = adUseClient
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    sql = "Select Top 1 nValFijo, nValFijoDia, nValVent, nValComp From TipoCambio WHERE  " _
        & "DateDiff(day,dFecCamb,'" & Format$(dFecha, gsFormatoFecha) & "') = 0 Order by dFecCamb Desc"
    Set rs = oCon.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        Select Case nTpoTipoCambio
            Case TCFijoMes
                EmiteTipoCambio = rs("nValFijo")
            Case TCFijoDia
                EmiteTipoCambio = rs("nValFijoDia")
            Case TCVenta
                EmiteTipoCambio = rs("nValVent")
            Case TCCompra
                EmiteTipoCambio = rs("nValComp")
        End Select
    End If
    rs.Close
    Set rs = Nothing
End Function


Public Function LeeConstSistema(ByVal sConstante As String) As String
    Dim rsVar As Recordset
    Dim sSql As String
    sSql = "SELECT nConsSisDesc, nConsSisValor FROM ConstSistema WHERE nConsSisCod = " & sConstante
    
    Set rsVar = New Recordset
    rsVar.CursorLocation = adUseClient
    rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsVar.ActiveConnection = Nothing
    If rsVar.EOF And rsVar.BOF Then
        LeeConstSistema = ""
    Else
        LeeConstSistema = rsVar("nConsSisValor")
    End If
    rsVar.Close
    Set rsVar = Nothing
End Function
Public Function GetCuentaNueva(ByVal sCuentaAnt As String) As String
Dim rsVar As Recordset
Dim sSql As String
Dim oConect As DConecta
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

sSql = "SELECT cCtaCod FROM RelCuentas Where cCtaCodAnt = '" & sCuentaAnt & "'"

Set rsVar = New Recordset
Set rsVar = oConect.CargaRecordSet(sSql)
If Not (rsVar.EOF And rsVar.BOF) Then
    GetCuentaNueva = rsVar("cCtaCod")
Else
    GetCuentaNueva = ""
End If
Set rsVar = Nothing
oConect.CierraConexion
Set oConect = Nothing
End Function

Public Function GetCuentaAntigua(ByVal sCuentaNue As String) As String
Dim rsVar As Recordset
Dim sSql As String
Dim oConect As DConecta
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

sSql = "SELECT cCtaCodAnt FROM RelCuentas Where cCtaCod = '" & sCuentaNue & "'"

Set rsVar = New Recordset
Set rsVar = oConect.CargaRecordSet(sSql)
If Not (rsVar.EOF And rsVar.BOF) Then
    GetCuentaAntigua = rsVar("cCtaCodAnt")
Else
    GetCuentaAntigua = ""
End If
Set rsVar = Nothing
oConect.CierraConexion
Set oConect = Nothing
End Function

Public Function GetMigroAsientoAgencia(ByVal pdFecha As Date, ByVal pnMoneda As Integer, gsCodAge As String, pdFecsis As Date) As Boolean
Dim rsAge As New ADODB.Recordset
Dim sSql As String
Dim gcMovNro As String
Dim oCon As DConecta
Set oCon = New DConecta
     
    'Saco de una agencia si se migro en la  moneda
    
    gcMovNro = Format(pdFecha, "yyyymmdd") & String(6, "0") & "112" & Right(gsCodAge, 2) & "00XXX" & pnMoneda
    
    sSql = "SELECT substring(cMovNro, 18,2) From Mov "
    sSql = sSql & " WHERE cMovNro like '" & Mid(gcMovNro, 1, 19) & "__" & Right(gcMovNro, 4) & "' and nMovEstado = '10' and nMovFlag <> '1'"
     
Set rsAge = dbCmact.CargaRecordSet(sSql)
If rsAge.BOF Then
    rsAge.Close
    Set rsAge = Nothing
    
    If Format(pdFecsis, "YYYYMMdd") = Format(pdFecha, "YYYYmmdd") Then
        sSql = "SELECT COUNT(nMovNro) as nCantidad From movdiario "
        sSql = sSql & "where Convert(varchar(8), cMovNro,112)='" & Format(pdFecha, "yyyymmdd") & "' "
        sSql = sSql & " and copeCod not in ('269994','900100') "
    
    Else
        sSql = "SELECT COUNT(nMovNro) as nCantidad From movdiario "
        sSql = sSql & " where Convert(varchar(8), cMovNro,112)='" & Format(pdFecha, "yyyymmdd") & "' "
        sSql = sSql & " and cOpeCod not in ('269994','900100') "
    End If
    
    
    oCon.AbreConexion 'Remota gsCodAge
    
    Set rsAge = oCon.CargaRecordSet(sSql)
    If rsAge.BOF Then
        GetMigroAsientoAgencia = False
    Else
        If rsAge!nCantidad > 0 Then
            GetMigroAsientoAgencia = False
        Else
            GetMigroAsientoAgencia = True
        End If
    End If
    oCon.CierraConexion
    
Else
    GetMigroAsientoAgencia = True
End If
rsAge.Close
Set rsAge = Nothing
  
End Function


'ARCV 11-09-2006
Public Function GetOperacionesUsuarioMatriz(ByVal sUsuario As String, ByVal MatFiltros As Variant, _
                                    ByVal vMatOpe As Variant, ByVal nNroRegOpe As Integer, Optional nMoneda As Moneda = 0, _
                                    Optional ByVal oConex As DConecta = Nothing) As ADODB.Recordset

Dim rsVar As New ADODB.Recordset
Dim sSql As String
Dim nCont As Integer
Dim sCodOpe As String
Dim i As Integer
Dim sFiltroMon As String

If nMoneda > 0 Then
    sFiltroMon = " AND O.cOpeCod NOT like '__" & Trim(nMoneda) & "%'"
End If

sSql = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O " _
    & "WHERE O.cOpeCod LIKE '" & MatFiltros(1) & "%' And O.cOpeVisible = '1'  " & sFiltroMon _
    & " "


For i = 2 To UBound(MatFiltros)
    sSql = sSql & " UNION " & _
           " SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O " _
           & "WHERE O.cOpeCod LIKE '" & MatFiltros(i) & "%' And O.cOpeVisible = '1'  " & sFiltroMon _
            & " "

Next i
sSql = sSql & " Order by O.cOpeCod, O.nOpeNiv"

'rsVar.CursorLocation = adUseClient
'rsVar.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'Set rsVar.ActiveConnection = Nothing
If oConex Is Nothing Then
    Set rsVar = dbCmact.CargaRecordSet(sSql)
Else
    Set rsVar = oConex.CargaRecordSet(sSql)
End If

nCont = 0

Do While Not rsVar.EOF
    For i = nCont To nNroRegOpe - 1
        If Trim(rsVar("cOpeCod")) = Trim(vMatOpe(i, 0)) Then
            sCodOpe = sCodOpe & Trim(rsVar("cOpeCod")) & "','"
            nCont = i
            Exit For
        End If
    Next i
    rsVar.MoveNext
Loop
If sCodOpe <> "" Then
    sCodOpe = Mid(sCodOpe, 1, Len(sCodOpe) - 2)
End If

sSql = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O "
If sCodOpe <> "" Then
    sSql = sSql & " Where O.cOpeCod in ('" & sCodOpe & ")"
Else
    sSql = sSql & " Where O.cOpeCod in ('')"
End If
sSql = sSql & " Order by O.cOpeCod, O.nOpeNiv"
rsVar.Close
'rsVar.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
If oConex Is Nothing Then
    Set rsVar = dbCmact.CargaRecordSet(sSql)
Else
    Set rsVar = oConex.CargaRecordSet(sSql)
End If
'Set rsVar.ActiveConnection = Nothing
Set GetOperacionesUsuarioMatriz = rsVar
Set rsVar = Nothing
End Function

Public Function GetOperacionesUsuario_NEW(ByVal sProducto As String, _
        Optional nMoneda As Moneda = 0, Optional ByVal prsOpeRep As ADODB.Recordset) As ADODB.Recordset

Dim sSql As String
Dim sFiltroMon As String
Dim rsTemp As ADODB.Recordset

Dim RSTEMP2 As ADODB.Recordset

If nMoneda > 0 Then
    'sFiltroMon = " cOpeCod NOT like '__" & Trim(nMoneda) & "%'"
    'sFiltroMon = " SUBSTRING(cOpeCod,3,1)<> " & Trim(nMoneda)
    'sFiltroMon = " cOpeCod NOT like '__" & Trim(nMoneda) & "%' AND "
    'sFiltroMon = " AND cOpeCod LIKE '__" & IIf(nMoneda = gMonedaExtranjera, gMonedaNacional, gMonedaExtranjera) & "%'"
    'sFiltroMon = "cOpeCod LIKE '__" & IIf(nMoneda = gMonedaExtranjera, gMonedaNacional, gMonedaExtranjera) & "%'"
    sProducto = Replace(sProducto, "M", IIf(nMoneda = gMonedaExtranjera, gMonedaNacional, gMonedaExtranjera))
    sProducto = sProducto & " OR cOpeCod LIKE '" & Replace(sProducto, IIf(nMoneda = gMonedaExtranjera, gMonedaNacional, gMonedaExtranjera), "0")
    sProducto = Mid(sProducto, 1, Len(sProducto) - 2)
End If

prsOpeRep.MoveFirst
Set rsTemp = prsOpeRep.Clone


'If sFiltroMon = "" Then
    rsTemp.Filter = "(cOpeCod LIKE '" & sProducto & "%')" '& sFiltroMon
    Set GetOperacionesUsuario_NEW = rsTemp
'Else
'    RSTEMP.Filter = "(cOpeCod LIKE '" & sProducto & "%')" '& sFiltroMon
    'Set RSTEMP2 = RSTEMP.Clone
'    RecordSetAdiciona RSTEMP2, RSTEMP
    
'    RSTEMP2.Filter = sFiltroMon
'    Set GetOperacionesUsuario_NEW = RSTEMP2
'End If
'If sFiltroMon <> "" Then
'    RSTEMP.Filter = sFiltroMon
'End If
'RSTEMP.Filter = sFiltroMon & " cOpeCod LIKE '" & sProducto & "%'"

Set rsTemp = Nothing
End Function

'RIRO20141106 ERS159 ***********************
Public Function GetConstante_New(ByVal nConstante As ConstanteCabecera, Optional sFiltro As String = "", _
Optional sFiltro2 As String = "", Optional sFiltro3 As String = "") As ADODB.Recordset
    Dim sSql As String, sFiltroCad As String, sFiltroCad2 As String, sFiltroCad3 As String
    Dim rsRel As New ADODB.Recordset
    'rsRel.CursorLocation = adUseClient
    If sFiltro <> "" Then sFiltroCad = " And nConsValor NOT IN (" & sFiltro & ") "
    If sFiltro2 <> "" Then sFiltroCad2 = " And nConsValor LIKE " & sFiltro2 & " "
    If sFiltro3 <> "" Then sFiltroCad3 = " And bEstado = 1 "
    sSql = "SELECT UPPER(cConsDescripcion) cDescripcion, nConsValor FROM " & sDBComunes & "Constante WHERE " _
        & "nConsCod = " & nConstante & " AND nConsCod <> nConsValor " & sFiltroCad & sFiltroCad2 & sFiltroCad3 & " ORDER BY nConsValor"
    'rsRel.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set rsRel = dbCmact.CargaRecordSet(sSql)
    
    'Set rsRel.ActiveConnection = Nothing
    Set GetConstante_New = rsRel
    Set rsRel = Nothing
End Function
'END RIRO ***********************************

Public Function GetTipoRelInst() As ADODB.Recordset
    Dim oConec As New DConecta
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    oConec.AbreConexion
    sSql = "Exec stp_sel_GetTipoRelInst"
    Set GetTipoRelInst = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function 'NAGL ERS074-2017 20171209
