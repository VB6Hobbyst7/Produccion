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
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long

Private dbCmact As Connection
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String


Public Function LeeConstSistema(ByVal sConstante As String) As String
Dim rsVar As ADODB.Recordset
Dim sSql As String
sSql = "SELECT nConsSisDesc, nConsSisValor FROM ConstSistema WHERE nConsSisCod = " & sConstante

Set rsVar = New ADODB.Recordset
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


Public Function GetNombreComputadora() As String
Dim sbuffMaq As String
Dim nSizeMaq As Long
sbuffMaq = Space(255)
nSizeMaq = Len(sbuffMaq)
GetComputerName sbuffMaq, nSizeMaq
GetNombreComputadora = Trim(Left$(sbuffMaq, nSizeMaq))
End Function

Public Function EmiteTipoCambio(ByVal dFecha As Date, ByVal nTpoTipoCambio As TipoCambio) As Double
Dim sql As String
Dim rs As ADODB.Recordset
EmiteTipoCambio = 0
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
sql = "Select nValFijo, nValFijoDia, nValVent, nValComp From TipoCambio WHERE  " _
    & "DateDiff(day,dFecCamb,'" & Format$(dFecha, "mm/dd/yyyy") & "') = 0"
rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
If Not rs.EOF And Not rs.BOF Then
    Select Case nTpoTipoCambio
        Case TCFijoMes
            EmiteTipoCambio = rs("nValFijo")
        Case TCFijoDia
            EmiteTipoCambio = rs("nValFijoDia")
        Case TCVenta
            EmiteTipoCambio = rs("nValVent")
        Case TCVenta
            EmiteTipoCambio = rs("nValComp")
    End Select
End If
rs.Close
Set rs = Nothing
End Function

Public Function GetPersonaEdad(ByVal sPersona As String) As Long
Dim rsPers As ADODB.Recordset
Dim sSql As String
sSql = "SELECT DATEDIFF(yy,dPersNacCreac,Getdate()) Edad FROM " & sDBPersona & "Persona " _
    & "WHERE cPersCod = '" & sPersona & "'"
Set rsPers = New ADODB.Recordset
rsPers.CursorLocation = adUseClient
rsPers.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsPers.ActiveConnection = Nothing
If rsPers.EOF And rsPers.BOF Then
    GetPersonaEdad = 0
Else
    GetPersonaEdad = rsPers("Edad")
End If
rsPers.Close
Set rsPers = Nothing
End Function

Public Function GetOperacionesUsuario(ByVal sUsuario As String, ByVal sProducto As String, _
        ByVal vMatOpe As Variant, ByVal nNroRegOpe As Integer, Optional nMoneda As Moneda = 0) As ADODB.Recordset

Dim rsVar As ADODB.Recordset
Dim sSql As String
Dim NCONT As Integer
Dim sCodOpe As String
Dim i As Integer
Dim sFiltroMon As String

If nMoneda > 0 Then
    sFiltroMon = " AND O.cOpeCod NOT like '__" & Trim(nMoneda) & "%'"
End If

sSql = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O " _
    & "WHERE O.cOpeCod LIKE '" & sProducto & "%' And O.cOpeVisible = '1'  " & sFiltroMon _
    & " Order by O.cOpeCod, O.nOpeNiv"

Set rsVar = New ADODB.Recordset
rsVar.CursorLocation = adUseClient
rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsVar.ActiveConnection = Nothing
NCONT = 0
Do While Not rsVar.EOF
    For i = NCONT To nNroRegOpe - 1
        If Trim(rsVar("cOpeCod")) = Trim(vMatOpe(i, 0)) Then
            sCodOpe = sCodOpe & Trim(rsVar("cOpeCod")) & "','"
            NCONT = i
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
rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsVar.ActiveConnection = Nothing
Set GetOperacionesUsuario = rsVar
Set rsVar = Nothing
End Function

Public Function GetOperacionesUsuario_NEW(ByVal sProducto As String, _
        Optional nMoneda As Moneda = 0, Optional ByVal prsOpeRep As ADODB.Recordset) As ADODB.Recordset

Dim sSql As String
Dim sFiltroMon As String
Dim RSTEMP As ADODB.Recordset

If nMoneda > 0 Then
    sFiltroMon = " AND cOpeCod NOT like '__" & Trim(nMoneda) & "%'"
End If

prsOpeRep.MoveFirst
Set RSTEMP = prsOpeRep.Clone

'RSTEMP.Find "cOpeCod LIKE '" & sProducto & "%'" & sFiltroMon
RSTEMP.Filter = "cOpeCod LIKE '" & sProducto & "%'" & sFiltroMon

Set GetOperacionesUsuario_NEW = RSTEMP
Set RSTEMP = Nothing
End Function

Public Function GetTrabajadorCMACT() As ADODB.Recordset
Dim sSql As String
Dim rsTrab As ADODB.Recordset
Set rsTrab = New ADODB.Recordset
rsTrab.CursorLocation = adUseClient
sSql = "Select P.cPersNombre, P.cPersCod FROM " & sDBPersona & "Persona P INNER JOIN RRHH T ON " _
    & "P.cPersCod = T.cPersCod Where nRHEstado NOT IN (" & gRHEstDespedido & "," _
    & gRHEstRetirado & ") ORDER BY cPersNombre"
rsTrab.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTrab.ActiveConnection = Nothing
Set GetTrabajadorCMACT = rsTrab
Set rsTrab = Nothing
End Function

Public Function GeneraNuevaCuenta(ByVal sAgencia As String, ByVal nProducto As Producto, ByVal nMoneda As Moneda) As String
Dim rsCta As ADODB.Recordset
Dim sCuenta As String, sCorr As String, sCorrDig
Dim sSql As String, sDigOK As String

sSql = "Select ISNULL(CONVERT(INT,MAX(SUBSTRING(cCtaCod,10,8))),0) + 1 Correlativo FROM " _
    & "Producto Where cCtaCod LIKE '%" & Trim(sAgencia) & Trim(nProducto) & Trim(nMoneda) & "%'"
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
If Not (rsCta.EOF And rsCta.BOF) Then
    sCorrDig = FillNum(Trim(rsCta("Correlativo")), 8, "0")
    sCorr = Trim(nProducto) & Trim(nMoneda) & FillNum(Trim(rsCta("Correlativo")), 8, "0")
    sDigOK = GetDigitoChequeo(sCorrDig)
    sCuenta = sAgencia & sCorr & sDigOK
End If
rsCta.Close
Set rsCta = Nothing
GeneraNuevaCuenta = sCuenta
End Function

Public Function GetDigitoChequeo(ByVal psCadena As String) As Integer
Dim liFactor As Variant
Dim liCadena(1 To 8) As Integer
Dim liSum, i As Integer
Dim lnDigito As Integer
liFactor = Array(9, 8, 7, 6, 5, 4, 3, 2)
liCadena(1) = Val(Mid(psCadena, 1, 1))
liCadena(2) = Val(Mid(psCadena, 2, 1))
liCadena(3) = Val(Mid(psCadena, 3, 1))
liCadena(4) = Val(Mid(psCadena, 4, 1))
liCadena(5) = Val(Mid(psCadena, 5, 1))
liCadena(6) = Val(Mid(psCadena, 6, 1))
liCadena(7) = Val(Mid(psCadena, 7, 1))
liCadena(8) = Val(Mid(psCadena, 8, 1))
liSum = 0
For i = 1 To 8
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

Public Function GetConstante(ByVal nConstante As ConstanteCabecera, Optional sFiltro As String = "", _
Optional sFiltro2 As String = "") As ADODB.Recordset
Dim sSql As String, sFiltroCad As String, sFiltroCad2 As String
Dim rsRel As ADODB.Recordset
Set rsRel = New ADODB.Recordset
rsRel.CursorLocation = adUseClient
If sFiltro <> "" Then sFiltroCad = " And nConsValor NOT IN (" & sFiltro & ") "
If sFiltro2 <> "" Then sFiltroCad2 = " And nConsValor LIKE " & sFiltro2 & " "
sSql = "SELECT UPPER(cConsDescripcion) cDescripcion, nConsValor FROM " & sDBComunes & "Constante WHERE " _
    & "nConsCod = " & nConstante & " AND nConsCod <> nConsValor " & sFiltroCad & sFiltroCad2 & " ORDER BY nConsValor"
rsRel.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsRel.ActiveConnection = Nothing
Set GetConstante = rsRel
Set rsRel = Nothing
End Function

Public Function GetDataUser(ByVal psCodDato As String, Optional pbBuscaXCodPers As Boolean = False) As ADODB.Recordset
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oConect As DConecta
    Set oConect = New DConecta
    Set rs = New ADODB.Recordset
    If oConect.AbreConexion = False Then Exit Function
    
    sql = " Select  P.cPersNombre , P.cPersDireccDomicilio  , RH.cPersCod,  Rh.cAreaCod , A.cAreaDescripcion, " _
        & "         cAgenciaActual, ISNULL(Ag.cAgeDescripcion,'') as cDescAgActual , RH.cAreaCodActual, Isnull(A1.cAreaDescripcion,'') as cDescAreaActual , " _
        & "         A.cAreaEstruc,  cAgenciaAsig, ISNULL(AG1.cAgeDescripcion,'') as cDescAgAsig , " _
        & "         ISNULL((SELECT cPersIDnro FROM PERSID WHERE cPersIDTpo ='" & gPersIdDNI & "' AND cPersCod = p.cPersCod ),'') AS DNI, " _
        & "         ISNULL((SELECT cPersIDnro FROM PERSID WHERE cPersIDTpo ='" & gPersIdRUC & "' AND cPersCod = p.cPersCod ),'') AS RUC,   " _
        & "         ISNULL(C.cRHCargoCod,'') AS cRHCargoCod,  ISNULL(C.cRHCargoDescripcion,'') AS cRHCargoDescripcion, " _
        & "         ISNULL(C.cConsDescripcion,'') AS cConsDescripcion, ISNULL(C.cRHCargoCategoria,'') cRHCargoCategoria , ISNULL(A.nLogTramite,0) AS  cLogTramite " _
        & " From    RRHH RH JOIN Persona P on P.cPersCod = RH.cPersCod " _
        & "         LEFT JOIN (SELECT MAX(dRHCargoFecha) as FechaCargo,C.cPersCod, C.cRHCargoCod, CT.cRHCargoDescripcion , CC.cConsDescripcion, CT.cRHCargoCategoria " _
        & "                    FROM     RHCargos C " _
        & "                             JOIN RHCargosTabla CT   ON CT.cRHCargoCod = C.cRHCargoCod " _
        & "                             JOIN CONSTANTE CC   ON CC.nConsValor  = CT.cRHCargoCategoria " _
        & "                     WHERE  CC.nCONSCOD LIKE '" & gRHCategoria & "' " _
        & "                    GROUP BY C.cPersCod,C.cRHCargoCod,CT.cRHCargoDescripcion , CC.cConsDescripcion, CT.cRHCargoCategoria ) as C " _
        & "         ON C.cPersCod = RH.cPersCod " _
        & "         LEFT JOIN Areas A ON A.cAreaCod = RH.cAreaCod " _
        & "         LEFT JOIN Agencias AG ON  AG.cAgeCod = LTRIM(RTRIM(RH.cAgenciaActual)) " _
        & "         LEFT JOIN Agencias AG1 on AG1.cAgecod= LTRIM(RTRIM(RH.cAgenciaAsig )) " _
        & "         LEFT JOIN Areas A1 ON A1.cAreaCod = RH.cAreaCodActual " _
        & " Where   " & IIf(pbBuscaXCodPers = True, " RH.cPersCod ='" & psCodDato & "' ", " cUser ='" & psCodDato & "' ")
        
    Set rs = oConect.CargaRecordSet(sql)
    Set GetDataUser = rs
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Public Function GetDatosPersona(ByVal psPersCod As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Set oConect = New DConecta
Set rs = New ADODB.Recordset
If oConect.AbreConexion = False Then Exit Function

sql = " Select  P.cPersNombre , P.cPersDireccDomicilio  , RH.cPersCod,  Rh.cAreaCod , A.cAreaDescripcion, " _
    & "         cAgenciaActual, ISNULL(Ag.cAgeDescripcion,'') as cDescAgActual , " _
    & "         A.cAreaEstruc,  cAgenciaAsig, ISNULL(AG1.cAgeDescripcion,'') as cDescAgAsig , " _
    & "         ISNULL((SELECT cPersIDnro FROM PERSID WHERE cPersIDTpo ='" & gPersIdDNI & "' AND cPersCod = p.cPersCod ),'') AS DNI, " _
    & "         ISNULL((SELECT cPersIDnro FROM PERSID WHERE cPersIDTpo ='" & gPersIdRUC & "' AND cPersCod = p.cPersCod ),'') AS RUC ,   " _
    & "         ISNULL(C.cRHCargoCod,'') AS cRHCargoCod ,  ISNULL(C.cRHCargoDescripcion,'')  AS cRHCargoDescripcion " _
    & " From    RRHH RH JOIN Persona P  ON P.cPersCod = RH.cPersCod " _
    & "         LEFT JOIN (SELECT MAX(dRHCargoFecha) as FechaCargo,C.cPersCod, C.cRHCargoCod, CT.cRHCargoDescripcion " _
    & "                    FROM     RHCargos C " _
    & "                             JOIN RHCargosTabla CT   ON CT.cRHCargoCod = C.cRHCargoCod " _
    & "                    GROUP BY C.cPersCod,C.cRHCargoCod,CT.cRHCargoDescripcion ) as C " _
    & "         ON C.cPersCod = RH.cPersCod " _
    & "         LEFT JOIN Areas A ON A.cAreaCod = RH.cAreaCod " _
    & "         LEFT JOIN Agencias AG ON  AG.cAgeCod = LTRIM(RTRIM(RH.cAgenciaActual)) " _
    & "         LEFT JOIN Agencias AG1 on AG1.cAgecod= LTRIM(RTRIM(RH.cAgenciaAsig )) " _
    & " Where   RH.cPersCod ='" & psPersCod & "' "
    
Set rs = oConect.CargaRecordSet(sql)
Set GetDatosPersona = rs
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

Public Function GetMensajeBoletas(ByVal sCuenta As String, Optional sAgencia As String) As String
Dim sSql As String
Dim rsMsg As ADODB.Recordset
Set rsMsg = New ADODB.Recordset
rsMsg.CursorLocation = adUseClient
sSql = "Select cMensaje from Mensajes where cCodPro = '" & Mid(sCuenta, 6, 3) & "' And bMenEst = 1"
If sAgencia <> "" Then
    sSql = sSql & " And cCodAge = '" & sAgencia & "'"
End If

rsMsg.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMsg.ActiveConnection = Nothing
If rsMsg.EOF And rsMsg.BOF Then
    GetMensajeBoletas = ""
Else
    GetMensajeBoletas = Trim(rsMsg("cMensaje"))
End If
rsMsg.Close
Set rsMsg = Nothing
End Function


Private Sub Class_Initialize()
Dim sConn As String
Dim ClsIni As COMConecta.DCOMClasIni
Dim oConecta As DConecta

Set ClsIni = New COMConecta.DCOMClasIni
sConn = ClsIni.CadenaConexion
sDBComunes = ClsIni.BaseComunes
sDBPersona = ClsIni.BasePersonas
sDBImagenes = ClsIni.BaseImagenes
Set ClsIni = Nothing
Set oConecta = New DConecta
oConecta.AbreConexion
Set dbCmact = oConecta.ConexionActiva
End Sub

Private Sub Class_Terminate()
dbCmact.Close
Set dbCmact = Nothing
End Sub

Public Function GetParametro(ByVal pnProdParam As Long, ByVal nParametro As Long) As Double
    Dim rsVar As ADODB.Recordset
    Dim sSql As String
    Dim oConect As DConecta
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then Exit Function
    
    sSql = "SELECT nParValor FROM Parametro WHERE nParCod = " & nParametro & " " _
        & "And nParProd = " & pnProdParam
    
    GetParametro = 0
    Set rsVar = New ADODB.Recordset
    Set rsVar = oConect.CargaRecordSet(sSql)
    If Not rsVar.EOF And Not rsVar.BOF Then
        GetParametro = rsVar("nParValor")
    End If
    rsVar.Close
    Set rsVar = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Public Function GetNombreAgencias(Optional Filtro1 As String = "") As ADODB.Recordset
    Dim rsVar As ADODB.Recordset
    Dim sSql As String, sSqlaux As String
    Dim oConect As DConecta
    Set oConect = New DConecta
    
    If Filtro1 <> "" Then
        sSqlaux = " and cagecod<>'" & Filtro1 & "'"
    End If
    
    If oConect.AbreConexion = False Then Exit Function
    
    sSql = " SELECT cAgeCod,cAgeDescripcion, 2 as Nivel, cAgeDireccion, cAgeTelefono, cSubCtaCod, "
    sSql = sSql & " nAgeEspecial, cUbiGeoCod, cUltimaActualizacion FROM Agencias where ((nrp is not null) or nrp>0 )" & sSqlaux
    sSql = sSql & " Order by cAgeDescripcion "
        
        
    Set rsVar = New ADODB.Recordset
    Set rsVar = oConect.CargaRecordSet(sSql)
    Set GetNombreAgencias = rsVar
    Set rsVar = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Public Function GetPerifericosPC(ByVal sNomPC As String) As ADODB.Recordset
    Dim rsVar As ADODB.Recordset
    Dim sSql As String
    Dim oConect As DConecta
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then Exit Function
    
    sSql = "SELECT P.nPeriferico, P.nPuerto, K.cConsDescripcion cPeriferico, K1.cConsDescripcion cPuerto, P.nMarca From " _
        & "Periferico P INNER JOIN " & sDBComunes & "Constante K ON P.nPeriferico = K.nConsValor INNER JOIN " _
        & sDBComunes & "Constante K1 ON P.nPuerto = K1.nConsValor Where P.cPCNombre = '" & sNomPC & "' " _
        & "And K.nConsCod = " & gTipoPeriferico & " And K1.nConsCod = " & gTipoPuertoSerial
    
    Set rsVar = New ADODB.Recordset
    Set rsVar = oConect.CargaRecordSet(sSql)
    Set GetPerifericosPC = rsVar
    Set rsVar = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Public Function EliminaPerifericosPC(ByVal sNomPC As String)
    Dim sSql As String
    Dim oConect As DConecta
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then Exit Function
    
    sSql = "Delete Periferico Where cPCNombre = '" & sNomPC & "'"
    oConect.Ejecutar sSql
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Public Function AgergaPerifericoPC(ByVal sNomPC As String, ByVal nPeriferico As TipoPeriferico, _
        ByVal nPuerto As TipoPuertoSerial, ByVal nMarca As TipoPinPad)
        
    Dim sSql As String
    Dim oConect As DConecta
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then Exit Function
    
    sSql = "Insert Periferico (cPCNombre,nPeriferico,nPuerto,nMarca) " _
        & "Values ('" & sNomPC & "'," & nPeriferico & "," & nPuerto & "," & nMarca & ")"
    oConect.Ejecutar sSql
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Public Function GetPuertoPeriferico(ByVal nPeriferico As TipoPeriferico) As TipoPuertoSerial

    Dim rsVar As ADODB.Recordset
    Dim sSql As String, sMaquina As String, buffMaq As String
    Dim oConect As DConecta
    Dim lSizeMaq As Long
    
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then Exit Function
    
    buffMaq = Space(255)
    lSizeMaq = Len(buffMaq)
    Call GetComputerName(buffMaq, lSizeMaq)
    sMaquina = Trim(Left$(buffMaq, lSizeMaq))
    
    sSql = "SELECT nPuerto From Periferico Where cPCNombre = '" & sMaquina & "' And " _
        & "nPeriferico = " & nPeriferico
    
    Set rsVar = New ADODB.Recordset
    Set rsVar = oConect.CargaRecordSet(sSql)
    If rsVar.EOF And rsVar.BOF Then
        GetPuertoPeriferico = -1
    Else
        GetPuertoPeriferico = rsVar("nPuerto")
    End If
    Set rsVar = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Public Function TienePermisoEspecial(ByVal sGrupo As String, ByVal nPermiso As CapPermisosEspeciales) As Boolean
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConect As DConecta
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then Exit Function
    
    sSql = "Select cName From Permiso Where cGrupoUsu = '" & sGrupo & "' And cName = '" & Trim(nPermiso) & "' And cMenuOpe = '2'"
    Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        TienePermisoEspecial = True
    Else
        TienePermisoEspecial = False
    End If
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Public Function GetPermisoEspecialUsuario(ByVal nPermiso As CapPermisosEspeciales, ByVal sUsuario As String, _
                Optional sDominio As String = "") As Boolean
Dim clsAcc As UAcceso
Dim sGruposTotal As String, sGrupo As String
Dim bExito As Boolean

Set clsAcc = New UAcceso
clsAcc.CargaGruposUsuario sUsuario, sDominio
bExito = False
sGrupo = clsAcc.DameGrupoUsuario
Do While sGrupo <> ""
    If TienePermisoEspecial(sGrupo, nPermiso) Then
        bExito = True
        Exit Do
    End If
    sGrupo = clsAcc.DameGrupoUsuario
Loop
Set clsAcc = Nothing
GetPermisoEspecialUsuario = bExito
End Function

Public Function GetCuentaNueva(ByVal sCuentaAnt As String) As String
Dim rsVar As ADODB.Recordset
Dim sSql As String
Dim oConect As DConecta
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

sSql = "SELECT cCtaCod FROM RelCuentas Where cCtaCodAnt = '" & sCuentaAnt & "'"

Set rsVar = New ADODB.Recordset
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
Dim rsVar As ADODB.Recordset
Dim sSql As String
Dim oConect As DConecta
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

sSql = "SELECT cCtaCodAnt FROM RelCuentas Where cCtaCod = '" & sCuentaNue & "'"

Set rsVar = New ADODB.Recordset
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

Public Function ValidaTarjeta(ByVal sTarjeta As String, ByVal sClave As String) As ResultVerificacionTarjeta


    Dim sPass As String, sSql As String
    Dim rs As ADODB.Recordset
    Dim oConect As DConecta
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then Exit Function
    
    Set rs = New ADODB.Recordset
    
    sSql = "Select cClave From Tarjeta Where cTarjCod = '" & sTarjeta & "'"
    
    Set rs = oConect.CargaRecordSet(sSql)
    
    If rs.EOF Then      'Faltaba verificacion
        ValidaTarjeta = gTarjNoRegistrada
        rs.Close
        Set rs = Nothing
        Set oConect = Nothing
        Exit Function
    End If
    
    sPass = Encripta(Trim(rs("cClave")), False)
    rs.Close
    Set rs = Nothing
    Set oConect = Nothing
    If sClave = sPass Then
        ValidaTarjeta = gClaveValida
    Else
        ValidaTarjeta = gClaveNOValida
    End If
    
End Function

Public Function GetNumDiasFeriado(ByVal dIni As Date, ByVal dFin As Date) As Integer
Dim nDias As Integer
Dim rsFer As ADODB.Recordset
Dim oConect As DConecta
Dim sSql As String

Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

sSql = "Select Count(Dia) nNum From (" _
    & "Select Convert(Varchar(10),dFeriado,102) Dia From Feriado Where DateDiff(dd,'" & Format$(dIni, "mm/dd/yyyy") & "',dFeriado) >= 0 " _
    & "And DateDiff(dd,dFeriado,'" & Format$(dFin, "mm/dd/yyyy") & "') >= 0 Group by Convert(Varchar(10),dFeriado,102)" _
    & ") A"
Set rsFer = oConect.CargaRecordSet(sSql)
Set rsFer.ActiveConnection = Nothing
If rsFer.EOF And rsFer.BOF Then
    GetNumDiasFeriado = 0
Else
    GetNumDiasFeriado = rsFer("nNum")
End If
rsFer.Close
Set rsFer = Nothing
oConect.CierraConexion
Set oConect = Nothing
End Function

Public Function GetAgencias() As ADODB.Recordset
Dim oConect As DConecta
Dim sSql As String

Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

sSql = "Select cAgeCod,cAgeDescripcion From Agencias Order by cAgeCod"
Set GetAgencias = oConect.CargaRecordSet(sSql)
oConect.CierraConexion
Set oConect = Nothing
End Function
Public Function GetAgenciasArbol() As ADODB.Recordset
Dim oConect As DConecta
Dim sSql As String

Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

sSql = "Select cAgeCod,cAgeDescripcion, 1 as nNivel From Agencias Order by cAgeCod"
Set GetAgenciasArbol = oConect.CargaRecordSet(sSql)
oConect.CierraConexion
Set oConect = Nothing
End Function
Public Function GetUserAreaAgencia(ByVal psAreaCod As String, Optional ByVal psAgecod As String, Optional psUser As String = "", Optional ByVal pbExcluir As Boolean = True) As ADODB.Recordset
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
    & " WHERE   RH.cAreaCodActual = '" & psAreaCod & "' AND RH.cAgenciaActual = '" & psAgecod & "' " _
    & "         AND RH.cUser<>'' " & lsFiltro _
    & " Order BY  P.cPersNombre"

Set rs = oConect.CargaRecordSet(sql)
Set GetUserAreaAgencia = rs
oConect.CierraConexion
Set oConect = Nothing
End Function

Public Function GetUserAreaAgenciaResumenIngEgre(ByVal psAreaCod As String, Optional ByVal psAgecod As String, Optional psUser As String = "", Optional ByVal pbExcluir As Boolean = True) As ADODB.Recordset
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

sql = " select * from " _
    & "     (SELECT     UPPER(RH.cUser) AS cUser, REPLACE(REPLACE(Replace(P.cPersNombre,'/',' '),'-',''),'\' ,' DE ') AS cNomUser , 1 Nivel  " _
    & "      FROM       RRHH RH " _
    & "                 JOIN PERSONA P ON P.CPERSCOD = RH.CPERSCOD " _
    & "      WHERE      RH.cAreaCodActual = '" & psAreaCod & "' AND RH.cAgenciaActual = '" & psAgecod & "' " _
    & "                 AND RH.cUser<>'' " & lsFiltro _
    & "       UNION     " _
    & "       SELECT    TOP 1 'BOVE' AS cUser, 'USUARIO DE BOVEDA' AS cNomUser,  1 AS Nivel FROM RRHH) AS CUSER  " _
    & " Order BY  cNomUser"

Set rs = oConect.CargaRecordSet(sql)
Set GetUserAreaAgenciaResumenIngEgre = rs
oConect.CierraConexion
Set oConect = Nothing
End Function
Public Function GetConstanteNiv(ByVal nConstante As ConstanteCabecera) As ADODB.Recordset
Dim sSql As String
Dim rsRel As ADODB.Recordset
Set rsRel = New ADODB.Recordset
rsRel.CursorLocation = adUseClient
sSql = "SELECT nConsValor , UPPER(cConsDescripcion) cDescripcion, 1 as Nivel FROM " & sDBComunes & "Constante WHERE " _
    & "nCONSCOD = " & nConstante & " AND nCONSCOD <> nConsValor ORDER BY nConsValor"
rsRel.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsRel.ActiveConnection = Nothing
Set GetConstanteNiv = rsRel
Set rsRel = Nothing
End Function

Public Function GetConstanteArbol(ByVal nConstante As ConstanteCabecera, Optional nModulo As Integer = 100) As ADODB.Recordset
Dim sSql As String
Dim rsRel As ADODB.Recordset
Set rsRel = New ADODB.Recordset
rsRel.CursorLocation = adUseClient
sSql = "SELECT nConsValor , UPPER(cConsDescripcion) cDescripcion, " _
    & "Nivel = CASE WHEN nConsValor%" & nModulo & " > 0 THEN 2 ELSE 1 END FROM " & sDBComunes & "Constante WHERE " _
    & "nCONSCOD = " & nConstante & " AND nCONSCOD <> nConsValor ORDER BY nConsValor"
rsRel.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsRel.ActiveConnection = Nothing
Set GetConstanteArbol = rsRel
Set rsRel = Nothing
End Function

Public Function GetCierreDiaRealizado(pdFecha As Date) As Boolean
Dim sSql As String
Dim rsRel As ADODB.Recordset
Set rsRel = New ADODB.Recordset
rsRel.CursorLocation = adUseClient
sSql = " Select Count(*) NFecha From ConstSistema Where nConsSisCod In (7,11,12,13) " _
     & " And Right(RTrim(LTrim(nConsSisValor)),4) +  Substring(RTrim(LTrim(nConsSisValor)),4,2) + Left(RTrim(LTrim(nConsSisValor)),2) <> '" & Format(pdFecha, gsFormatoMovFecha) & "'"
rsRel.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText


If rsRel.EOF And rsRel.BOF Then
    GetCierreDiaRealizado = False
Else
    If rsRel!NFecha = 4 Then
        GetCierreDiaRealizado = False
    Else
        GetCierreDiaRealizado = True
    End If
End If

Set rsRel.ActiveConnection = Nothing
Set rsRel = Nothing
End Function


