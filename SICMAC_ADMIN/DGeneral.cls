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
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long


Private dbCmact As Connection
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String

Public Function GetCOM(ByVal sDispositivo As String) As String
    Dim rsPC As Recordset
    Dim sSql As String
    Dim sMaquina As String
    
    sMaquina = GetNombreComputadora
    sSql = "Select cPuerto from SetupCOM where cDispositivo = '" & sDispositivo & "' " _
        & "and cNomPC = '" & sMaquina & "'"
    
    Set rsPC = New Recordset
    rsPC.CursorLocation = adUseClient
    rsPC.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsPC.ActiveConnection = Nothing
    If rsPC.EOF And rsPC.BOF Then
        GetCOM = ""
    Else
        GetCOM = "COM" & Trim(rsPC!cPuerto)
    End If
    rsPC.Close
    Set rsPC = Nothing
End Function

Public Function LeeVarSistema(ByVal sProd As String, ByVal sVariable As String) As String
    Dim rsVar As Recordset
    Dim sSql As String
    sSql = "SELECT cValorVar, cDescVar FROM VarSistema WHERE cCodProd = '" & sProd & "' " _
        & "AND cNomVar = '" & sVariable & "'"
    Set rsVar = New Recordset
    rsVar.CursorLocation = adUseClient
    rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsVar.ActiveConnection = Nothing
    If rsVar.EOF And rsVar.BOF Then
        LeeVarSistema = ""
    Else
        LeeVarSistema = rsVar("cValorVar")
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

Public Function GetPersonaEdad(ByVal sPersona As String) As Long
    Dim rsPers As Recordset
    Dim sSql As String
    sSql = "SELECT DATEDIFF(yy,dPersNacCreac,Getdate()) Edad FROM " & sDBPersona & "Persona " _
        & "WHERE cPersCod = '" & sPersona & "'"
    Set rsPers = New Recordset
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

Public Function GetOperacionesUsuario(ByVal sUsuario As String, ByVal sProducto As String, Optional ByVal psMoneda As String = "") As Recordset
    Dim rsVar As Recordset
    Dim sSql As String
    Dim lsFiltroMon As String
    If psMoneda <> "" Then
        lsFiltroMon = " AND O.cOpeCod NOT like '__" & psMoneda & "%'"
    End If
    sSql = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O " _
        & "WHERE O.cOpeCod LIKE '" & sProducto & "%' And O.cOpeVisible = '1'  " & lsFiltroMon _
        & " Order by O.cOpeCod, O.nOpeNiv"
    Set rsVar = New Recordset
    rsVar.CursorLocation = adUseClient
    rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsVar.ActiveConnection = Nothing
    Set GetOperacionesUsuario = rsVar
    Set rsVar = Nothing
  
End Function

Public Function GetTrabajadorCMACT() As Recordset
    Dim sSql As String
    Dim rsTrab As Recordset
    Set rsTrab = New Recordset
    rsTrab.CursorLocation = adUseClient
    sSql = "Select P.cPersNombre, P.cPersCod FROM " & sDBPersona & "Persona P INNER JOIN RRHH T ON " _
        & "P.cPersCod = T.cPersCod Where cRHEstado NOT IN (" & gRHEstDespedido & "," _
        & gRHEstRetirado & ") ORDER BY cPersNombre"
    rsTrab.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsTrab.ActiveConnection = Nothing
    Set GetTrabajadorCMACT = rsTrab
    Set rsTrab = Nothing
End Function

Public Function GeneraNuevaCuenta(ByVal sAgencia As String, ByVal nProducto As Producto, ByVal nMoneda As Moneda) As String
    Dim rsCta As Recordset
    Dim sCuenta As String, sCorr As String, sCorrDig
    Dim sSql As String, sDigOK As String
    
    sSql = "Select ISNULL(CONVERT(INT,MAX(SUBSTRING(cCtaCod,10,8))),0) + 1 Correlativo FROM " _
        & "Producto Where cCtaCod LIKE '%" & Trim(sAgencia) & Trim(nProducto) & Trim(nMoneda) & "%'"
    Set rsCta = New Recordset
    rsCta.CursorLocation = adUseClient
    rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsCta.ActiveConnection = Nothing
    If Not (rsCta.EOF And rsCta.BOF) Then
        sCorrDig = Trim(nProducto) & Trim(nMoneda) & Right(FillNum(Trim(rsCta("Correlativo")), 8, "0"), 5)
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
    Dim liCadena(1 To 12) As Integer
    Dim liSum, i As Integer
    Dim lnDigito As Integer
    liFactor = Array(10, 9, 8, 7, 6, 5, 4, 3, 2)
    liCadena(1) = Val(Mid(psCadena, 1, 1))
    liCadena(2) = Val(Mid(psCadena, 2, 1))
    liCadena(3) = Val(Mid(psCadena, 3, 1))
    liCadena(4) = Val(Mid(psCadena, 4, 1))
    liCadena(5) = Val(Mid(psCadena, 5, 1))
    liCadena(6) = Val(Mid(psCadena, 6, 1))
    liCadena(7) = Val(Mid(psCadena, 7, 1))
    liCadena(8) = Val(Mid(psCadena, 8, 1))
    liCadena(9) = Val(Mid(psCadena, 9, 1))
    liSum = 0
    For i = 1 To 9
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
    Set rsRel = New Recordset
    rsRel.CursorLocation = adUseClient
    sSql = "SELECT UPPER(cConsDescripcion) cDescripcion, nConsValor FROM " & sDBComunes & "Constante WHERE " _
        & "nConsCod = " & nConstante & " AND nConsCod <> nConsValor ORDER BY nConsValor"
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

Public Function GetMensajeBoletas(ByVal sCuenta As String) As String
    Dim sSql As String
    Dim rsMsg As Recordset
    Set rsMsg = New Recordset
    rsMsg.CursorLocation = adUseClient
    sSql = "Select cMensaje from Mensajes where cCodPro = '" & Mid(sCuenta, 6, 3) & "'"
    rsMsg.Open sSql, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
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
    Dim ClsIni As ClsIni.ClasIni
    Set ClsIni = New ClsIni.ClasIni
    sConn = ClsIni.CadenaConexion
    sDBComunes = ClsIni.BaseComunes
    sDBPersona = ClsIni.BasePersonas
    sDBImagenes = ClsIni.BaseImagenes
    Set ClsIni = Nothing
    Set dbCmact = New Connection
    dbCmact.Open sConn
End Sub

Private Sub Class_Terminate()
    dbCmact.Close
    Set dbCmact = Nothing
End Sub

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



