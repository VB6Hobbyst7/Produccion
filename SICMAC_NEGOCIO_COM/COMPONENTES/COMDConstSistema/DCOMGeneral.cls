VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMGeneral"
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
Dim rsVar As New ADODB.Recordset
Dim sSql As String
sSql = "SELECT nConsSisDesc, nConsSisValor FROM ConstSistema WHERE nConsSisCod = " & sConstante


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

'AMDO TI-ERS086-2013 20130726 ****************************************************
'RIRO20141106ERS159 sPersCod
Public Function VerificaExistePermisoCargo(ByVal psCodCargo As String, ByVal pnTipoOpc As Integer, _
                                        Optional sPersCod As String = "") As Boolean
Dim rsVar As New ADODB.Recordset
Dim sSql As String
sSql = "stp_sel_ERS0862013_VerificaPermisoCargo '" & psCodCargo & "'," & pnTipoOpc & ", '" & sPersCod & "'" 'RIRO20141106ERS159 sPersCod


rsVar.CursorLocation = adUseClient
rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsVar.ActiveConnection = Nothing
If rsVar.EOF And rsVar.BOF Then
    'VerificaExistePermisoCargo = ""
    VerificaExistePermisoCargo = False 'RIRO20141217 Correccion
Else
    VerificaExistePermisoCargo = rsVar!bExiste
End If
rsVar.Close
Set rsVar = Nothing
End Function
'END AMDO

Public Function GetNombreComputadora() As String
Dim sbuffMaq As String
Dim nSizeMaq As Long
sbuffMaq = Space(255)
nSizeMaq = Len(sbuffMaq)
GetComputerName sbuffMaq, nSizeMaq
GetNombreComputadora = Trim(Left$(sbuffMaq, nSizeMaq))
End Function

Public Function EmiteTipoCambio(ByVal dFecha As Date, ByVal nTpoTipoCambio As TipoCambio) As Double
Dim Sql As String
Dim rs As New ADODB.Recordset
EmiteTipoCambio = 0

rs.CursorLocation = adUseClient
'FRHU 20140515 ERS068-2014 RQ14286: Se agrego nValPondVenta en el script
Sql = "Select nValFijo, nValFijoDia, nValVent, nValComp, nValPondVenta From TipoCambio WHERE  " _
    & "DateDiff(day,dFecCamb,'" & Format$(dFecha, "mm/dd/yyyy") & "') = 0"
rs.Open Sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
If Not rs.EOF And Not rs.BOF Then
    Select Case nTpoTipoCambio
        Case TCFijoMes
            EmiteTipoCambio = rs("nValFijo")
        Case TCFijoDia
            EmiteTipoCambio = rs("nValFijoDia")
        Case TCVenta
            EmiteTipoCambio = rs("nValVent")
        '***Modificado por ELRO el 20120716, según OYP-RFC024-2012
        'Case TCVenta
        Case TCCompra
        '***Fin Modificado por ELRO el 20120716*******************
            EmiteTipoCambio = rs("nValComp")
        Case TCPondVenta 'FRHU 20140515 ERS068-2014 RQ14286
            EmiteTipoCambio = rs("nValPondVenta")
    End Select
End If
rs.Close
Set rs = Nothing
End Function

Public Function GetPersonaEdad(ByVal sPersona As String) As Long
Dim rsPers As New ADODB.Recordset
Dim sSql As String
sSql = "SELECT DATEDIFF(yy,dPersNacCreac,Getdate()) Edad FROM " & sDBPersona & "Persona " _
    & "WHERE cPersCod = '" & sPersona & "'"

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
        ByVal vMatOpe As Variant, ByVal nNroRegOpe As Integer, Optional nmoneda As Moneda = 0) As ADODB.Recordset

Dim rsVar As New ADODB.Recordset
Dim sSql As String
Dim nCont As Integer
Dim sCodOpe As String
Dim i As Integer
Dim sFiltroMon As String

If nmoneda > 0 Then
    sFiltroMon = " AND O.cOpeCod NOT like '__" & Trim(nmoneda) & "%'"
End If

sSql = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O " _
    & "WHERE O.cOpeCod LIKE '" & sProducto & "%' And O.cOpeVisible = '1'  " & sFiltroMon _
    & " Order by O.cOpeCod, O.nOpeNiv"
rsVar.CursorLocation = adUseClient
rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsVar.ActiveConnection = Nothing
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
rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsVar.ActiveConnection = Nothing
Set GetOperacionesUsuario = rsVar
Set rsVar = Nothing
End Function

Public Function GetOperacionesUsuario_NEW(ByVal sProducto As String, _
        Optional nmoneda As Moneda = 0, Optional ByVal prsOpeRep As ADODB.Recordset) As ADODB.Recordset

Dim sSql As String
Dim sFiltroMon As String
Dim rsTemp As ADODB.Recordset

If nmoneda > 0 Then
    sFiltroMon = " AND O.cOpeCod NOT like '__" & Trim(nmoneda) & "%'"
End If

prsOpeRep.MoveFirst
Set rsTemp = prsOpeRep.Clone
rsTemp.Filter = "cOpeCod LIKE '" & sProducto & "%'" & sFiltroMon

Set GetOperacionesUsuario_NEW = rsTemp
Set rsTemp = Nothing
End Function


'Operaciones del Usuario Cargados con la Matriz de Filtros

Public Function GetOperacionesUsuarioMatriz(ByVal sUsuario As String, ByVal MatFiltros As Variant, _
                                    ByVal vMatOpe As Variant, ByVal nNroRegOpe As Integer, Optional nmoneda As Moneda = 0) As ADODB.Recordset

Dim rsVar As New ADODB.Recordset
Dim sSql As String
Dim nCont As Integer
Dim sCodOpe As String
Dim i As Integer
Dim sFiltroMon As String

If nmoneda > 0 Then
    sFiltroMon = " AND O.cOpeCod NOT like '__" & Trim(nmoneda) & "%'"
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

rsVar.CursorLocation = adUseClient
rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsVar.ActiveConnection = Nothing
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
rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsVar.ActiveConnection = Nothing
Set GetOperacionesUsuarioMatriz = rsVar
Set rsVar = Nothing
End Function

Public Function GetTrabajadorCMACT() As ADODB.Recordset
Dim sSql As String
Dim rsTrab As New ADODB.Recordset

rsTrab.CursorLocation = adUseClient
sSql = "Select P.cPersNombre, P.cPersCod FROM " & sDBPersona & "Persona P INNER JOIN RRHH T ON " _
    & "P.cPersCod = T.cPersCod Where nRHEstado NOT IN (" & gRHEstDespedido & "," _
    & gRHEstRetirado & ") ORDER BY cPersNombre"
rsTrab.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTrab.ActiveConnection = Nothing
Set GetTrabajadorCMACT = rsTrab
Set rsTrab = Nothing
End Function

Public Function FechaHora(ByVal psFecha As Date) As String
    Dim oConec As New COMConecta.DCOMConecta
    If oConec.AbreConexion = True Then
        FechaHora = Format(psFecha & Space(1) & oConec.GetHoraServer, "mm/dd/yyyy hh:mm:ss")
    Else
        FechaHora = ""
    End If
    Set oConec = Nothing
End Function

Public Function GeneraNuevaCuenta(ByVal sAgencia As String, ByVal nProducto As Producto, ByVal nmoneda As Moneda) As String
    Dim rsCta As New ADODB.Recordset
    Dim sCuenta As String, sCorr As String, sCorrDig
    Dim sSql As String, sDigOK As String
    Dim oFun As New COMFunciones.FCOMCadenas


    sSql = "Select ISNULL(CONVERT(INT,MAX(SUBSTRING(cCtaCod,10,8))),0) + 1 Correlativo FROM " _
        & "Producto Where cCtaCod LIKE '%" & Trim(sAgencia) & Trim(nProducto) & Trim(nmoneda) & "%'"

    rsCta.CursorLocation = adUseClient
    rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsCta.ActiveConnection = Nothing
    If Not (rsCta.EOF And rsCta.BOF) Then
        sCorrDig = oFun.FillNum(Trim(rsCta("Correlativo")), 8, "0")
        sCorr = Trim(nProducto) & Trim(nmoneda) & oFun.FillNum(Trim(rsCta("Correlativo")), 8, "0")
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
'MADM 20110630 - sFiltroCad3 As String
Public Function GetConstante(ByVal nConstante As ConstanteCabecera, Optional sFiltro As String = "", _
Optional sFiltro2 As String = "", Optional sFiltro3 As String = "") As ADODB.Recordset
    Dim sSql As String, sFiltroCad As String, sFiltroCad2 As String, sFiltroCad3 As String
    Dim rsRel As New ADODB.Recordset
    rsRel.CursorLocation = adUseClient
    If sFiltro <> "" Then sFiltroCad = " And nConsValor NOT IN (" & sFiltro & ") "
    If sFiltro2 <> "" Then sFiltroCad2 = " And nConsValor LIKE " & sFiltro2 & " "
    If sFiltro3 <> "" Then sFiltroCad3 = " And bEstado = 1 "
    sSql = "SELECT UPPER(cConsDescripcion) cDescripcion, nConsValor FROM " & sDBComunes & "Constante WHERE " _
        & "nConsCod = " & nConstante & " AND nConsCod <> nConsValor " & sFiltroCad & sFiltroCad2 & sFiltroCad3 & " ORDER BY nConsValor"
    rsRel.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsRel.ActiveConnection = Nothing
    Set GetConstante = rsRel
    Set rsRel = Nothing
End Function

Public Function GetDataUser(ByVal psCodDato As String, Optional pbBuscaXCodPers As Boolean = False) As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
   
    If oCon.AbreConexion = False Then Exit Function
    ''FRHU20170919 INCIDENTE INC1709150014
    'modificado por gitu 07/05/2008
    'EJVG20120821 Adecuado para obtener Cargo RRHH
    ''Sql = " Select  P.cPersNombre , P.cPersDireccDomicilio  , RH.cPersCod,  Rh.cAreaCod , A.cAreaDescripcion, " _
        & "         cAgenciaActual, ISNULL(Ag.cAgeDescripcion,'') as cDescAgActual , RH.cAreaCodActual, Isnull(A1.cAreaDescripcion,'') as cDescAreaActual , " _
        & "         A.cAreaEstruc,  cAgenciaAsig, ISNULL(AG1.cAgeDescripcion,'') as cDescAgAsig , " _
        & "         ISNULL((SELECT cPersIDnro FROM PERSID WHERE cPersIDTpo ='" & gPersIdDNI & "' AND cPersCod = p.cPersCod ),'') AS DNI, " _
        & "         ISNULL((SELECT cPersIDnro FROM PERSID WHERE cPersIDTpo ='" & gPersIdRUC & "' AND cPersCod = p.cPersCod ),'') AS RUC,   " _
        & "         ISNULL(C.cRHCargoCod,'') AS cRHCargoCod,  ISNULL(C.cRHCargoDescripcion,'') AS cRHCargoDescripcion, " _
        & "         ISNULL(C.cConsDescripcion,'') AS cConsDescripcion, ISNULL(C.cRHCargoCategoria,'') cRHCargoCategoria , ISNULL(A.nLogTramite,0) AS  cLogTramite, FechaCargo " _
        & " From    RRHH RH JOIN Persona P on P.cPersCod = RH.cPersCod " _
        & "         LEFT JOIN (SELECT dRHCargoFecha as FechaCargo,C.cPersCod, C.cRHCargoCod, CT.cRHCargoDescripcion , CC.cConsDescripcion, CT.cRHCargoCategoria " _
        & "                    FROM     RHCargos C " _
        & "                             JOIN RHCargosTabla CT   ON CT.cRHCargoCod = C.cRHCargoCod " _
        & "                             LEFT JOIN CONSTANTE CC   ON CC.nConsValor  = CT.cRHCargoCategoria And CC.nCONSCOD LIKE '" & gRHCategoria & "' " _
        & "                     WHERE  C.dRHCargoFecha = (Select MAX(dRHCargoFecha) From RHCargos Where cPerscod = C.cPerscod)) as C " _
        & "         ON C.cPersCod = RH.cPersCod " _
        & "         LEFT JOIN Areas A ON A.cAreaCod = RH.cAreaCod " _
        & "         LEFT JOIN Agencias AG ON  AG.cAgeCod = LTRIM(RTRIM(RH.cAgenciaActual)) " _
        & "         LEFT JOIN Agencias AG1 on AG1.cAgecod= LTRIM(RTRIM(RH.cAgenciaAsig )) " _
        & "         LEFT JOIN Areas A1 ON A1.cAreaCod = RH.cAreaCodActual " _
        & " Where   " & IIf(pbBuscaXCodPers = True, " RH.cPersCod ='" & psCodDato & "' ", " cUser ='" & psCodDato & "' ")
        'GROUP BY C.cPersCod,C.cRHCargoCod,CT.cRHCargoDescripcion , CC.cConsDescripcion, CT.cRHCargoCategoria
    Sql = "exec stp_sel_GetDataUser '" & psCodDato & "'"
    ''FRHU20170919 INCIDENTE
    Set rs = oCon.CargaRecordSet(Sql)
    Set GetDataUser = rs
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GetDatosPersona(ByVal psPersCod As String) As ADODB.Recordset
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function

Sql = " Select  P.cPersNombre , P.cPersDireccDomicilio  , RH.cPersCod,  Rh.cAreaCod , A.cAreaDescripcion, " _
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
    
Set rs = oCon.CargaRecordSet(Sql)
Set GetDatosPersona = rs
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function CuentaNombre(ByVal pCodCta As String) As String
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    Dim oCon As COMConecta.DCOMConecta
    
    Set oCon = New COMConecta.DCOMConecta
    If oCon.AbreConexion = False Then Exit Function
    
    pCodCta = Trim(pCodCta)
    If Len(pCodCta) > 4 Then
       tmpSql = "SELECT RTRIM(ISNULL(c1.cCtaContDesc,'')) + ' ' + RTRIM(c.cCtaContDesc) as cCtaContDesc " _
              & "FROM   " & sDBComunes & "CtaCont c LEFT JOIN " & sDBComunes & "CtaCont c1 ON c1.cCtaContCod = substring(c.cCtaContCod,1," & Len(pCodCta) - 2 & ") " _
              & "WHERE  c.cCtaContCod = '" & pCodCta & "'"
    Else
       tmpSql = "SELECT cCtaContDesc FROM " & sDBComunes & "CtaCont WHERE cCtaContCod = '" & pCodCta & "'"
    End If
    Set tmpReg = oCon.CargaRecordSet(tmpSql)
    If (tmpReg.BOF Or tmpReg.EOF) Then
        CuentaNombre = ""
    Else
        CuentaNombre = Trim(tmpReg!cCtaContDesc)
    End If
    tmpReg.Close: Set tmpReg = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GetMensajeBoletas(ByVal sCuenta As String, Optional sAgencia As String) As String
Dim sSql As String
Dim rsMsg As New ADODB.Recordset

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
    Dim ocona As New COMConecta.DCOMConecta
    If ocona.AbreConexion = True Then
       Set dbCmact = ocona.ConexionActiva
    End If
End Sub

Private Sub Class_Terminate()
dbCmact.Close
Set dbCmact = Nothing
End Sub

Public Function GetParametro(ByVal pnProdParam As Long, ByVal nParametro As Long) As Double
    Dim rsVar As New ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    If oCon.AbreConexion = False Then Exit Function
    
    sSql = "SELECT nParValor FROM Parametro WHERE nParCod = " & nParametro & " " _
        & "And nParProd = " & pnProdParam
    
    GetParametro = 0
    Set rsVar = oCon.CargaRecordSet(sSql)
    If Not rsVar.EOF And Not rsVar.BOF Then
        GetParametro = rsVar("nParValor")
    End If
    rsVar.Close
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GetNombreAgencias(Optional Filtro1 As String = "") As ADODB.Recordset
    Dim rsVar As New ADODB.Recordset
    Dim sSql As String, sSqlaux As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    'RECO20140621 ERS095-2014***************************************
    'If Filtro1 <> "" Then
    '    sSqlaux = " and cagecod<>'" & Filtro1 & "'"
    'End If
    
    If oCon.AbreConexion = False Then Exit Function
    
    'sSql = " SELECT cAgeCod,cAgeDescripcion, 2 as Nivel, cAgeDireccion, cAgeTelefono, cSubCtaCod, "
    'sSql = sSql & " nAgeEspecial, cUbiGeoCod, cUltimaActualizacion FROM Agencias where ((nrp is not null) or nrp>0 )" & sSqlaux
    'sSql = sSql & " Order by cAgeDescripcion "
    sSql = "stp_sel_DevuelveListaAgencias " & Filtro1
    'RECO FIN*******************************************************
    Set rsVar = oCon.CargaRecordSet(sSql)
    Set GetNombreAgencias = rsVar
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GetPerifericosPC(ByVal sNomPC As String) As ADODB.Recordset
    Dim rsVar As New ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    If oCon.AbreConexion = False Then Exit Function
    
    sSql = "SELECT P.nPeriferico, P.nPuerto, K.cConsDescripcion cPeriferico, K1.cConsDescripcion cPuerto, P.nMarca From " _
        & "Periferico P INNER JOIN " & sDBComunes & "Constante K ON P.nPeriferico = K.nConsValor INNER JOIN " _
        & sDBComunes & "Constante K1 ON P.nPuerto = K1.nConsValor Where P.cPCNombre = '" & sNomPC & "' " _
        & "And K.nConsCod = " & gTipoPeriferico & " And K1.nConsCod = " & gTipoPuertoSerial
    
    Set rsVar = oCon.CargaRecordSet(sSql)
    Set GetPerifericosPC = rsVar
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function EliminaPerifericosPC(ByVal sNomPC As String)
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    If oCon.AbreConexion = False Then Exit Function
    
    sSql = "Delete Periferico Where cPCNombre = '" & sNomPC & "'"
    oCon.Ejecutar sSql
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function AgergaPerifericoPC(ByVal sNomPC As String, ByVal nPeriferico As TipoPeriferico, _
        ByVal nPuerto As TipoPuertoSerial, ByVal nMarca As TipoPinPad)
        
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    If oCon.AbreConexion = False Then Exit Function
    
    sSql = "Insert Periferico (cPCNombre,nPeriferico,nPuerto,nMarca) " _
        & "Values ('" & sNomPC & "'," & nPeriferico & "," & nPuerto & "," & nMarca & ")"
    oCon.Ejecutar sSql
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GetPuertoPeriferico(ByVal nPeriferico As TipoPeriferico, ByVal psMaquina As String) As TipoPuertoSerial

    Dim rsVar As New ADODB.Recordset
    Dim sSql As String, sMaquina As String, buffMaq As String
    Dim oCon As COMConecta.DCOMConecta
    Dim lSizeMaq As Long
    
    Set oCon = New COMConecta.DCOMConecta
    If oCon.AbreConexion = False Then Exit Function
    
'    buffMaq = Space(255)
'    lSizeMaq = Len(buffMaq)
'    Call GetComputerName(buffMaq, lSizeMaq)
'    sMaquina = Trim(Left$(buffMaq, lSizeMaq))
    
    sSql = "SELECT nPuerto From Periferico Where cPCNombre = '" & psMaquina & "' And " _
        & "nPeriferico = " & nPeriferico
    
    Set rsVar = oCon.CargaRecordSet(sSql)
    If rsVar.EOF And rsVar.BOF Then
        GetPuertoPeriferico = -1
    Else
        GetPuertoPeriferico = rsVar("nPuerto")
    End If
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function TienePermisoEspecial(ByVal sGrupo As String, ByVal nPermiso As CapPermisosEspeciales) As Boolean
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    If oCon.AbreConexion = False Then Exit Function
    
    sSql = "Select cName From Permiso Where cGrupoUsu = '" & sGrupo & "' And cName = '" & Trim(nPermiso) & "' And cMenuOpe = '2'"
    Set rs = oCon.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        TienePermisoEspecial = True
    Else
        TienePermisoEspecial = False
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GetPermisoEspecialUsuario(ByVal nPermiso As CapPermisosEspeciales, ByVal sUsuario As String, _
                Optional sDominio As String = "") As Boolean
Dim clsAcc As DCOMUAcceso
Dim sGruposTotal As String, sGrupo As String
Dim bExito As Boolean

Set clsAcc = New DCOMUAcceso
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
Dim rsVar As New ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
If oCon.AbreConexion = False Then Exit Function

sSql = "SELECT cCtaCod FROM RelCuentas Where cCtaCodAnt = '" & sCuentaAnt & "'"

Set rsVar = oCon.CargaRecordSet(sSql)
If Not (rsVar.EOF And rsVar.BOF) Then
    GetCuentaNueva = rsVar("cCtaCod")
Else
    GetCuentaNueva = ""
End If
Set rsVar = Nothing
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function GetCuentaAntigua(ByVal sCuentaNue As String) As String
Dim rsVar As New ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
If oCon.AbreConexion = False Then Exit Function

sSql = "SELECT cCtaCodAnt FROM RelCuentas Where cCtaCod = '" & sCuentaNue & "'"
Set rsVar = oCon.CargaRecordSet(sSql)
If Not (rsVar.EOF And rsVar.BOF) Then
    GetCuentaAntigua = rsVar("cCtaCodAnt")
Else
    GetCuentaAntigua = ""
End If
Set rsVar = Nothing
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function ValidaTarjeta(ByVal sTarjeta As String, ByVal sClave As String) As ResultVerificacionTarjeta

    Dim sPass As String, sSql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Dim oCls As COMConecta.DCOMClasIni
    
    Set oCon = New COMConecta.DCOMConecta
    If oCon.AbreConexion = False Then Exit Function
    
    sSql = "Select cClave From Tarjeta Where cTarjCod = '" & sTarjeta & "'"
    
    Set rs = oCon.CargaRecordSet(sSql)
    
    If rs.EOF Then      'Faltaba verificacion
        ValidaTarjeta = gTarjNoRegistrada
        rs.Close
        Set rs = Nothing
        Set oCon = Nothing
        Exit Function
    End If
    Set oCls = New COMConecta.DCOMClasIni
    sPass = oCls.Encripta(Trim(rs("cClave")), False)
    rs.Close
    Set rs = Nothing
    Set oCon = Nothing
    If sClave = sPass Then
        ValidaTarjeta = gClaveValida
    Else
        ValidaTarjeta = gClaveNOValida
    End If
    
End Function

Public Function GetNumDiasFeriado(ByVal dIni As Date, ByVal dFin As Date) As Integer
Dim nDias As Integer
Dim rsFer As New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String

Set oCon = New COMConecta.DCOMConecta
If oCon.AbreConexion = False Then Exit Function

sSql = "Select Count(Dia) nNum From (" _
    & "Select Convert(Varchar(10),dFeriado,102) Dia From Feriado Where DateDiff(dd,'" & Format$(dIni, "mm/dd/yyyy") & "',dFeriado) >= 0 " _
    & "And DateDiff(dd,dFeriado,'" & Format$(dFin, "mm/dd/yyyy") & "') >= 0 Group by Convert(Varchar(10),dFeriado,102)" _
    & ") A"
Set rsFer = oCon.CargaRecordSet(sSql)
Set rsFer.ActiveConnection = Nothing
If rsFer.EOF And rsFer.BOF Then
    GetNumDiasFeriado = 0
Else
    GetNumDiasFeriado = rsFer("nNum")
End If
rsFer.Close
Set rsFer = Nothing
oCon.CierraConexion
Set oCon = Nothing
End Function
'MADM 20110307 - STORE
Public Function GetAgencias() As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String

Set oCon = New COMConecta.DCOMConecta
If oCon.AbreConexion = False Then Exit Function

    'sSql = "Select cAgeCod,cAgeDescripcion From Agencias Order by cAgeCod"
    sSql = "exec stp_sel_DevolverAgencias"

Set GetAgencias = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function
'MADM 20110307 - STORE
Public Function GetAgenciasArbol() As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String

Set oCon = New COMConecta.DCOMConecta
If oCon.AbreConexion = False Then Exit Function

'sSql = "Select cAgeCod,cAgeDescripcion, 1 as nNivel From Agencias Order by cAgeCod"
sSql = "exec stp_sel_DevolverAgenciasArbol"
Set GetAgenciasArbol = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetUserAreaAgencia(ByVal psAreaCod As String, Optional ByVal psAgeCod As String, Optional psUser As String = "", Optional ByVal pbExcluir As Boolean = True) As ADODB.Recordset
Dim Sql As String
Dim lsFiltro As String
Dim rs As New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsql As String, lcExcluir As String
Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function

'*** PEAC 20090822
If pbExcluir Then
    lcExcluir = "V"
Else
    lcExcluir = "F"
End If
'*** FIN PEAC

'*** PEAC 20090922
'If psUser <> "" Then
'    lsFiltro = " AND RH.cUser " & IIf(pbExcluir, " NOT ", "") & " IN ('" & psUser & "')"
'End If
'
'If psAreaCod = "067" Then
'    lsql = " WHERE   RH.cAreaCodActual in('" & psAreaCod & "','026') AND RH.cAgenciaActual = '" & psAgeCod & "' "
'Else
'    lsql = " WHERE   RH.cAreaCodActual in('" & psAreaCod & "') AND RH.cAgenciaActual = '" & psAgeCod & "' "
'End If
'
'Sql = "SELECT   UPPER(RH.cUser) AS cUser, REPLACE(REPLACE(Replace(P.cPersNombre,'/',' '),'-',''),'\' ,' DE ') AS cNomUser , 1 Nivel  " _
'    & " FROM    RRHH RH JOIN PERSONA P ON P.CPERSCOD = RH.CPERSCOD "
'Sql = Sql & lsql _
'    & "         AND RH.cUser<>'' " & lsFiltro _
'    & " Order BY  P.cPersNombre"
'*** FIN PEAC

'*** PEAC 20090922
psAreaCod = Replace(Replace(psAreaCod, "'", ""), " ", "")
psUser = Replace(Replace(psUser, "'", ""), " ", "")

Sql = " exec stp_sel_ObtieneUsuariosPorAreaAgencia '" & psAreaCod & "','" & psAgeCod & "','" & psUser & "','" & lcExcluir & "'"
'*** FIN PEAC

Set rs = oCon.CargaRecordSet(Sql)
Set GetUserAreaAgencia = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
'VAPA20161022
Public Function GetUserAgenciaCierre(ByVal psFechaSis As String) As ADODB.Recordset
Dim Sql As String
Dim lsFiltro As String
Dim rs As New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsql As String, lcExcluir As String
Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function


Sql = " stp_sel_ObtieneUsuariosCierre '" & Format(psFechaSis, "YYYYmmdd") & "'"


Set rs = oCon.CargaRecordSet(Sql)
Set GetUserAgenciaCierre = rs

oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetAgenciaxUser(ByVal psUser As String) As String

Dim oCon As COMConecta.DCOMConecta
Dim Sql As String
Set oCon = New COMConecta.DCOMConecta
Dim rs As New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function

Sql = " stp_sel_AgenciaxUser '" & psUser & "'"

Set rs = oCon.CargaRecordSet(Sql)
    GetAgenciaxUser = rs!Agencia

oCon.CierraConexion
Set oCon = Nothing
End Function



'FIN VAPA20161022
Public Function GetUserAreaAgenciaResumenIngEgre(ByVal psAreaCod As String, Optional ByVal psAgeCod As String, Optional psUser As String = "", Optional ByVal pbExcluir As Boolean = True) As ADODB.Recordset
Dim Sql As String
Dim lsFiltro As String
Dim rs As New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function

If psUser <> "" Then
    lsFiltro = " AND RH.cUser " & IIf(pbExcluir, " NOT ", "") & " IN ('" & psUser & "')"
End If

Sql = " select * from " _
    & "     (SELECT     UPPER(RH.cUser) AS cUser, REPLACE(REPLACE(Replace(P.cPersNombre,'/',' '),'-',''),'\' ,' DE ') AS cNomUser , 1 Nivel  " _
    & "      FROM       RRHH RH " _
    & "                 JOIN PERSONA P ON P.CPERSCOD = RH.CPERSCOD " _
    & "      WHERE      RH.cAreaCodActual = '" & psAreaCod & "' AND RH.cAgenciaActual = '" & psAgeCod & "' AND RH.nRHEstado =201" _
    & "                 AND RH.cUser<>'' " & lsFiltro _
    & "       UNION     " _
    & "       SELECT    TOP 1 'BOVE' AS cUser, 'USUARIO DE BOVEDA' AS cNomUser,  1 AS Nivel FROM RRHH) AS CUSER  " _
    & " Order BY  cNomUser"

Set rs = oCon.CargaRecordSet(Sql)
Set GetUserAreaAgenciaResumenIngEgre = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetConstanteNiv(ByVal nConstante As ConstanteCabecera) As ADODB.Recordset
Dim sSql As String
Dim rsRel As New ADODB.Recordset
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
Dim rsRel As New ADODB.Recordset

rsRel.CursorLocation = adUseClient
'sSql = "SELECT nConsValor , UPPER(cConsDescripcion) cDescripcion, " _
'    & "Nivel = CASE WHEN nConsValor%" & nModulo & " > 0 THEN 2 ELSE 1 END FROM " & sDBComunes & "Constante WHERE " _
'    & "nCONSCOD = " & nConstante & " AND nCONSCOD <> nConsValor ORDER BY nConsValor"
sSql = "SELECT nConsValor , UPPER(cConsDescripcion) cDescripcion, " _
    & "Nivel = CASE WHEN SubString(Cast(nConsValor as Varchar(5)), 2, 2) <> 00 THEN 2 ELSE 1 END FROM " & sDBComunes & "Constante WHERE " _
    & "nCONSCOD = " & nConstante & " AND nCONSCOD <> nConsValor ORDER BY nConsValor"
rsRel.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsRel.ActiveConnection = Nothing
Set GetConstanteArbol = rsRel
Set rsRel = Nothing
End Function

Public Function GetCierreDiaRealizado(pdFecha As Date) As Boolean
Dim sSql As String
Dim rsRel As New ADODB.Recordset
Dim oFun As New COMFunciones.FCOMVarPublicas

rsRel.CursorLocation = adUseClient
sSql = " Select Count(*) NFecha From ConstSistema Where nConsSisCod In (7,11,12,13) " _
     & " And Right(RTrim(LTrim(nConsSisValor)),4) +  Substring(RTrim(LTrim(nConsSisValor)),4,2) + Left(RTrim(LTrim(nConsSisValor)),2) <> '" & Format(pdFecha, oFun.gsFormatoMovFecha) & "'"
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

Public Function GetTipCambio(pdFecha As Date, pnTipoCambio As TipoCambio) As Currency
'Dim oDGeneral As nTipoCambio
'Set oDGeneral = New nTipoCambio
GetTipCambio = EmiteTipoCambio(pdFecha, pnTipoCambio)

End Function

Public Function GetFechaMov(cMovnro, lDia As Boolean) As String
Dim lFec As Date
Dim oFun As New COMFunciones.FCOMVarPublicas
lFec = Mid(cMovnro, 7, 2) & "/" & Mid(cMovnro, 5, 2) & "/" & Mid(cMovnro, 1, 4)
If lDia Then
   GetFechaMov = Format(lFec, oFun.gsFormatoFechaView)
Else
   GetFechaMov = Format(lFec, oFun.gsFormatoFecha)
End If
End Function

Public Function CargaCargos() As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo MensaError
    
    sSql = " Select crhcargodescripcion + space(150-len(crhcargodescripcion))+crhcargocod Cargo from rhcargostabla where len(crhcargocod)>3 "
    sSql = sSql & "order by crhcargodescripcion "
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    Set CargaCargos = rs
    Exit Function
        
MensaError:
    Err.Raise vbObjectError + 100, "CargaCargos", "Error emitir documento en Operacion"
End Function

Private Function CargaOperaciones() As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    On Error GoTo MensaError
    
    sSql = sSql & " Select f.copedesc +':'+ ch.copedesc + space(150-len(f.copedesc +':'+ ch.copedesc))+ ch.copecod as operacion from opetpo ch  "
    sSql = sSql & " inner join opetpo f on f.copecod=left(ch.copecod,2)+'0000' "
    sSql = sSql & " where ch.copecod like  '2%[1-9]00'  "
    sSql = sSql & " order by  f.copedesc,ch.copedesc,ch.copecod  "
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
 
    Set CargaOperaciones = rs
   
    Exit Function
        
        
MensaError:
    Err.Raise vbObjectError + 100, "CargaOperaciones", "Error al consultar Operaciones"

End Function
Private Function CargaOperaciones2003() As ADODB.Recordset

    Dim rs As ADODB.Recordset
    Dim sSql As String
   On Error GoTo MensaError
    sSql = "select cOpeDesc, cOpeCod from opetpo where left(cOpecod,4)='2003'"
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set CargaOperaciones2003 = rs
    
    Exit Function
MensaError:
    Err.Raise vbObjectError + 100, "CargaOperaciones", "Error al consultar Operaciones"
End Function

Public Function InicioDia(ByVal psCodUser As String, _
                            ByVal psCodAge As String, _
                            ByRef pdFecSis As Date, ByVal psMov As String)

'Dim cmd As New ADODB.Command
'Dim prm As New ADODB.Parameter
Dim sSql As String
Dim rsTip As New ADODB.Recordset
Dim PObjConec As COMConecta.DCOMConecta
    
    Set PObjConec = New COMConecta.DCOMConecta
    PObjConec.AbreConexion
    PObjConec.ConexionActiva.BeginTrans
    pdFecSis = pdFecSis + 1
    
    'Valorizacion de Cheques de Captaciones
    'ALPA 08/07/2008
''    cmd.CommandText = "CapValorizaCheque"
''    cmd.CommandType = adCmdStoredProc
''    cmd.Name = "CapValorizaCheque"
''    Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
''    cmd.Parameters.Append prm
''    Set prm = cmd.CreateParameter("Usuario", adChar, adParamInput, 4)
''    cmd.Parameters.Append prm
''    Set prm = cmd.CreateParameter("Agencia", adChar, adParamInput, 2)
''    cmd.Parameters.Append prm
''    Set cmd.ActiveConnection = PObjConec.ConexionActiva
''    cmd.CommandTimeout = 720
''    cmd.Parameters.Refresh
''    PObjConec.ConexionActiva.CapValorizaCheque Format(pdFecSis & " " & PObjConec.GetHoraServer(), "dd/mm/yyyy hh:mm:ss"), psCodUser, Right(psCodAge, 2)
'ALPA***************************************************************
    sSql = ""
    sSql = "exec CapValorizaCheque '" & Format(pdFecSis & " " & PObjConec.GetHoraServer(), "YYYY/MM/DD hh:mm:ss") & "', '" & psCodUser & "', '" & Right(psCodAge, 2) & "'"
    PObjConec.ConexionActiva.Execute sSql
'*******************************************************************
'    Set cmd = Nothing
'    Set prm = Nothing
    
    'Valorizacion de Cheques de Captaciones
    
'    Set cmd = New ADODB.Command
'    Set prm = New ADODB.Parameter
'
'    cmd.CommandText = "ColocValorizaCheque"
'    cmd.CommandType = adCmdStoredProc
'    cmd.Name = "ColocValorizaCheque"
'    Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
'    cmd.Parameters.Append prm
'    Set prm = cmd.CreateParameter("Usuario", adChar, adParamInput, 4)
'    cmd.Parameters.Append prm
'    Set prm = cmd.CreateParameter("Agencia", adChar, adParamInput, 2)
'    cmd.Parameters.Append prm
'    Set cmd.ActiveConnection = PObjConec.ConexionActiva
'    cmd.CommandTimeout = 720
'    cmd.Parameters.Refresh
'    PObjConec.ConexionActiva.ColocValorizaCheque Format(pdFecSis & " " & PObjConec.GetHoraServer(), "dd/mm/yyyy hh:mm:ss"), psCodUser, Right(psCodAge, 2)
    sSql = ""
    sSql = "exec ColocValorizaCheque '" & Format(pdFecSis & " " & PObjConec.GetHoraServer(), "YYYY/MM/DD hh:mm:ss") & "', '" & psCodUser & "', '" & Right(psCodAge, 2) & "'"
    PObjConec.ConexionActiva.Execute sSql
    
    'Set cmd = Nothing
    'Set prm = Nothing
    
    'ARCV 06-11-2006
    'ejecutamos el calculo de cobro de notificaciones
    'cmd.CommandText = "sp_ComisionNotificaciones"
    'cmd.CommandType = adCmdStoredProc
    'cmd.Name = "ComisionNotificaciones"
    'Set prm = cmd.CreateParameter("dFecCierre", adDate, adParamInput)
    'cmd.Parameters.Append prm
    'Set cmd.ActiveConnection = PObjConec.ConexionActiva
    'cmd.CommandTimeout = 720
    'cmd.Parameters.Refresh
    'PObjConec.ConexionActiva.ComisionNotificaciones Format(pdFecSis & " " & PObjConec.GetHoraServer(), "dd/mm/yyyy hh:mm:ss")
    '------------------------------------
    
    'Actualiza Fecha de Inicio de Dia
    sSql = " UPDATE ConstSistema SET nConsSisValor = '" & Format(pdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 15"
    PObjConec.ConexionActiva.Execute sSql
    'Actualiza Fecha de Inicio del Sistema
    sSql = " UPDATE ConstSistema SET nConsSisValor = '" & Format(pdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 16"
    PObjConec.ConexionActiva.Execute sSql
    
    'Actualiza Fecha de Inicio del Sistema
    sSql = " Delete MovDiario "
    PObjConec.ConexionActiva.Execute sSql
    
    sSql = "UPDATE ConstSistema set nConsSisValor =0 WHERE nConsSisCod = 2"
    PObjConec.ConexionActiva.Execute sSql
    
    'Inserat Tipo de Cambio
''    ALPA 20081027**/**********************************************************
''    sSql = "Select nValVent,nValComp,nValFijo,nValFijoDia,nValVentEsp,nValCompEsp,nValPond,nValPondVenta from TipoCambio " & _
''           "where dFecCamb = (select Max(dFecCamb) from TipoCambio)"
    'ALPA 20140318
     sSql = "Select nValVent,nValComp,nValFijo,nValFijoDia,nValVentEsp,nValCompEsp,nValPond,nValPondVenta,isnull(nValPondREU,0) nValPondREU,nValSBSDia=isnull(nValSBSDia,0),nValCompTr=isnull(nValCompTr,0),nValVentTr=isnull(nValVentTr,0)  from TipoCambio " & _
           "where dFecCamb = (select Max(dFecCamb) from TipoCambio)"

'    sSql = "Select nValVent,nValComp,nValFijo,nValFijoDia,nValVentEsp,nValCompEsp,nValPond,nValPondVenta,isnull(nValPondREU,0) nValPondREU from TipoCambio " & _
'           "where dFecCamb = (select Max(dFecCamb) from TipoCambio)"
''    ***********************************************************************
    Set rsTip = PObjConec.CargaRecordSet(sSql)
    
    If Not (rsTip.EOF And rsTip.BOF) Then
        Dim pnValVent As Double
        Dim pnValComp As Double
        Dim pnValFijo As Double
        Dim pnValFijoDia  As Double
        Dim pnValVentEsp As Double
        Dim pnValCompEsp As Double
        Dim pnValPond As Double
        Dim pnValPondVenta As Double
        'ALPA 20081027****************************************************************
        Dim pnValPondREU As Double
        '*****************************************************************************
        'ALPA 20140226****************************************************************
        Dim lnValSBSDia As Currency
        Dim lnValCompTr As Currency
        Dim lnValVentTr As Currency
        '*****************************************************************************
        pnValVent = rsTip!nValVent
        pnValComp = rsTip!nValComp
        pnValFijo = rsTip!nValFijo
        pnValFijoDia = rsTip!nValFijoDia
        pnValVentEsp = rsTip!nValVentEsp
        pnValCompEsp = rsTip!nValCompEsp
        pnValPond = rsTip!nValPond
        pnValPondVenta = rsTip!nValPondVenta
        'ALPA 20081027****************************************************************
        pnValPondREU = rsTip!nValPondREU
        '*****************************************************************************
        'ALPA 20140226****************************************************************
        lnValSBSDia = rsTip!nValSBSDia
        lnValCompTr = rsTip!nValCompTr
        lnValVentTr = rsTip!nValVentTr
        '*****************************************************************************
''      ALPA 20081027**************************************************
''        sSql = " Insert TipoCambio (dFecCamb,nValVent,nValComp,nValFijo,nValFijoDia,nValVentEsp,nValCompEsp,nValPond,nValPondVenta,cUltimaActualizacion)" & _
''               " Values('" & Format(pdFecSis, "yyyymmdd") & "'," & pnValVent & "," & pnValComp & "," & pnValFijo & "," & pnValFijoDia & "," & pnValVentEsp & "," & pnValCompEsp & "," & pnValPond & "," & pnValPondVenta & ",'" & psMov & "')"
'       ALPA 20140318
'        sSql = " Insert TipoCambio (dFecCamb,nValVent,nValComp,nValFijo,nValFijoDia,nValVentEsp,nValCompEsp,nValPond,nValPondVenta,cUltimaActualizacion,nValPondREU)" & _
'               " Values('" & Format(pdFecSis, "yyyymmdd") & "'," & pnValVent & "," & pnValComp & "," & pnValFijo & "," & pnValFijoDia & "," & pnValVentEsp & "," & pnValCompEsp & "," & pnValPond & "," & pnValPondVenta & ",'" & psMov & "'," & pnValPondREU & ")"
'        PObjConec.ConexionActiva.Execute sSql
        sSql = " Insert TipoCambio (dFecCamb,nValVent,nValComp,nValFijo,nValFijoDia,nValVentEsp,nValCompEsp,nValPond,nValPondVenta,cUltimaActualizacion,nValPondREU,nValSBSDia,nValCompTr,nValVentTr)" & _
               " Values('" & Format(pdFecSis, "yyyymmdd") & "'," & pnValVent & "," & pnValComp & "," & pnValFijo & "," & pnValFijoDia & "," & pnValVentEsp & "," & pnValCompEsp & "," & pnValPond & "," & pnValPondVenta & ",'" & psMov & "'," & pnValPondREU & ", " & lnValSBSDia & ", " & lnValCompTr & ", " & lnValVentTr & ")"
        PObjConec.ConexionActiva.Execute sSql

    End If
    'ALPA 20081128***********************************************************
    rsTip.Close
    Set rsTip = Nothing

    sSql = "exec stp_sel_devolverTipoCambioEspecial "
    Set rsTip = PObjConec.CargaRecordSet(sSql)
    
If Not rsTip.EOF And Not rsTip.BOF Then
        Dim pdFecha As Date
        Dim pnIdRango As Integer
        Dim pnIdRangoDet As Integer
        Dim pnCompra As Double
        Dim pnVenta  As Double
    Do While Not rsTip.EOF
        pnIdRango = rsTip!nIdRango
        pnIdRangoDet = rsTip!nIdRangoDet
        pnCompra = rsTip!nCompra
        pnVenta = rsTip!nVenta
        sSql = "exec stp_ins_TipoCambioEspecial '" & Format(pdFecSis, "MM/DD/YYYY") & "', " & pnIdRango & "," & pnIdRangoDet & "," & pnCompra & "," & pnVenta
        PObjConec.ConexionActiva.Execute sSql
        rsTip.MoveNext
    Loop
End If
    
    '***********************************************************************
    PObjConec.ConexionActiva.CommitTrans
    PObjConec.CierraConexion
    
    
    Set PObjConec = Nothing
   
End Function
'MADM 20110307 - STORE
Public Function CargaInstituciones() As ADODB.Recordset
Dim R As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ERRORCargaAnalistas
'    sSql = "select P.cPersNombre, PT.cPersCod "
'    sSql = sSql & " from PersTpo PT Inner join Persona P ON PT.cPersCod = P.cPersCod And PT.nPersTipo = 1"
'    sSql = sSql & " Order By P.cPersNombre"
    sSql = "exec stp_sel_DevolverInstituciones "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not R.EOF And Not R.BOF Then
        Set CargaInstituciones = R
    End If
    'R.Close
    Set R = Nothing
    Exit Function
ERRORCargaAnalistas:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Public Function CargaAnalistas() As ADODB.Recordset
Dim R As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sAnalistas As String


    On Error GoTo ERRORCargaAnalistas
    
    sAnalistas = LeeConstSistema(gConstSistRHCargoCodAnalistas)
   
    
    sSql = "Select R.cPersCod, P.cPersNombre from RRHH R inner join Persona P ON R.cPersCod = P.cpersCod "
    sSql = sSql & " AND nRHEstado = 201 "
    sSql = sSql & " inner join RHCargos RC ON R.cPersCod = RC.cPersCod "
    sSql = sSql & " where  RC.cRHCargoCod in (" & sAnalistas & ") AND RC.dRHCargoFecha = (select MAX(dRHCargoFecha) from RHCargos RHC2 where RHC2.cPersCod = RC.cPersCod) "
    sSql = sSql & " order by P.cPersNombre "
        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not R.EOF And Not R.BOF Then
        Set CargaAnalistas = R
    End If
    'R.Close
    Set R = Nothing
    Exit Function
ERRORCargaAnalistas:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

'peac 20070913
'MADM 20100307 - STORE
Public Function CargaCampanhas() As ADODB.Recordset
Dim R As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCampanhas As String

    On Error GoTo ERRORCargaCampanhas
    
'    sCampanhas = LeeConstSistema(gConstSistRHCargoCodAnalistas)
    
'    sSql = "Select R.cPersCod, P.cPersNombre from RRHH R inner join Persona P ON R.cPersCod = P.cpersCod "
'    sSql = sSql & " AND nRHEstado = 201 "
'    sSql = sSql & " inner join RHCargos RC ON R.cPersCod = RC.cPersCod "
'    sSql = sSql & " where  RC.cRHCargoCod in (" & sAnalistas & ") AND RC.dRHCargoFecha = (select MAX(dRHCargoFecha) from RHCargos RHC2 where RHC2.cPersCod = RC.cPersCod) "
'    sSql = sSql & " order by P.cPersNombre "

'    sSql = "SELECT cdescripcion,idcampana from CAMPANAS order by idcampana"
      
    sSql = "EXEC stp_sel_DevolverCampanhas"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not R.EOF And Not R.BOF Then
        Set CargaCampanhas = R
    End If
    'R.Close
    Set R = Nothing
    Exit Function
ERRORCargaCampanhas:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

'ALPA 20100707 B2***********************************************
Public Function GetConstanteArbolCreditoYAhorro(ByVal nConstante As ConstanteCabecera, Optional nModulo As Integer = 100, Optional ByVal nConstante2 As ConstanteCabecera = "0", Optional psSTAhorro As String = "0") As ADODB.Recordset
Dim sSql As String
Dim rsRel As New ADODB.Recordset

rsRel.CursorLocation = adUseClient
sSql = "SELECT nConsValor , UPPER(cConsDescripcion) cDescripcion, " _
    & "Nivel = CASE WHEN SubString(Cast(nConsValor as Varchar(5)), 2, 2) <> 00 THEN 2 ELSE 1 END FROM " & sDBComunes & "Constante WHERE " _
    & "nCONSCOD = " & nConstante & " AND nCONSCOD <> nConsValor "
If nConstante2 <> "0" Then
    sSql = sSql & " UNION SELECT nConsValor , UPPER(cConsDescripcion) cDescripcion, " _
        & "Nivel = CASE WHEN SubString(Cast(nConsValor as Varchar(5)), 2, 2) <> 30 THEN 2 ELSE 1 END FROM " & sDBComunes & "Constante WHERE " _
        & "nCONSCOD = " & nConstante2 & " AND SubString(Cast(nConsValor as Varchar(5)),1,2) in ('" & psSTAhorro & "') AND nCONSCOD <> nConsValor ORDER BY nConsValor"
End If
rsRel.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsRel.ActiveConnection = Nothing
Set GetConstanteArbolCreditoYAhorro = rsRel
Set rsRel = Nothing
End Function
'*****************************************************************

'MAVM 20100723 JG
Public Function GeneraNvoNumeroJoyGar() As String
    Dim rsCod As New ADODB.Recordset
    Dim sCodigo As String
    Dim sSql As String
    Dim oFun As New COMFunciones.FCOMCadenas

    sSql = "Select ISNULL(CONVERT(INT,MAX(cJoyGarCod)),0) + 1 Correlativo FROM JoyaGarantia"

    rsCod.CursorLocation = adUseClient
    rsCod.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsCod.ActiveConnection = Nothing
    If Not (rsCod.EOF And rsCod.BOF) Then
        sCodigo = oFun.FillNum(Trim(rsCod("Correlativo")), 8, "0")
    End If
    rsCod.Close
    Set rsCod = Nothing
    GeneraNvoNumeroJoyGar = sCodigo
End Function

'***Agregado por ELRO el 20120723, según OYP-RFC077-2012
Public Function devolverAgenciasParaCoberturar() As ADODB.Recordset
Dim oDCOMConecta As COMConecta.DCOMConecta
Dim rsAgencias As ADODB.Recordset
Dim lsSQL As String
    
    On Error GoTo devolverAgenciasParaCoberturarError
    
    lsSQL = "exec stp_sel_RFC0772012_DevolverAgenciasParaCoberturar"
    Set oDCOMConecta = New COMConecta.DCOMConecta
    oDCOMConecta.AbreConexion
    Set rsAgencias = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion
    
    
    If Not rsAgencias.EOF And Not rsAgencias.BOF Then
        Set devolverAgenciasParaCoberturar = rsAgencias
    End If

    Set rsAgencias = Nothing
    Set oDCOMConecta = Nothing
    Exit Function
devolverAgenciasParaCoberturarError:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Public Function registrarPeriodoCoberturaAgencia(ByVal pnMesDes As Integer, _
                                                 ByVal pnAnioDes As Integer, _
                                                 ByVal pnMesHas As Integer, _
                                                 ByVal pnAnioHas As Integer, _
                                                 ByVal psMovNro As String) As Integer
Dim oDCOMConecta As COMConecta.DCOMConecta
Dim rsPeriodo As ADODB.Recordset
Dim lsSQL As String
    
    On Error GoTo registrarPeriodoCoberturaAgenciaError
    
    lsSQL = "exec stp_ins_RFC0772012_RegistrarPeriodoCoberturaAgencia " & pnMesDes & ", " & pnAnioDes & ", " & pnMesHas & ", " & pnAnioHas & ",  '" & psMovNro & "'"
    
    Set oDCOMConecta = New COMConecta.DCOMConecta
    Set rsPeriodo = New ADODB.Recordset
    
    oDCOMConecta.AbreConexion
    
    Set rsPeriodo = oDCOMConecta.CargaRecordSet(lsSQL)
    
    If Not rsPeriodo.BOF And Not rsPeriodo.EOF Then
        registrarPeriodoCoberturaAgencia = rsPeriodo!IdCobAgePer
    Else
        registrarPeriodoCoberturaAgencia = 0
    End If
    
    oDCOMConecta.CierraConexion
        
    Set rsPeriodo = Nothing
    Set oDCOMConecta = Nothing
    Exit Function
registrarPeriodoCoberturaAgenciaError:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Public Function registrarCoberturaAgencia(ByVal psAgeCod As String, _
                                          ByVal pnTipoCobertura As Integer, _
                                          ByVal pnIdCobAgePer As Integer, _
                                          ByVal pnMoneda As Integer, _
                                          ByVal pnCobertura As Currency, _
                                          ByVal psMovNro As String) As Integer
Dim oDCOMConecta As COMConecta.DCOMConecta
Dim rsCobertura As ADODB.Recordset
Dim lsSQL As String
    
    On Error GoTo registrarCoberturaAgenciaError
    
    lsSQL = "exec stp_ins_RFC0772012_RegistrarCoberturaAgencia '" & psAgeCod & "', " & pnTipoCobertura & ", " & pnIdCobAgePer & ", " & pnMoneda & ", " & pnCobertura & ", '" & psMovNro & "'"
    Set oDCOMConecta = New COMConecta.DCOMConecta
    Set rsCobertura = New ADODB.Recordset
    
    oDCOMConecta.AbreConexion
    
    Set rsCobertura = oDCOMConecta.CargaRecordSet(lsSQL)
    
    If Not rsCobertura.BOF And Not rsCobertura.EOF Then
        registrarCoberturaAgencia = rsCobertura!IdCoberturaAgencia
    Else
        registrarCoberturaAgencia = 0
    End If
    
    oDCOMConecta.CierraConexion
        
    Set rsCobertura = Nothing
    Set oDCOMConecta = Nothing
    Exit Function
registrarCoberturaAgenciaError:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Public Sub eliminarCoberturasAgencias(ByVal pnIdCobAgePer As Integer)
Dim oDCOMConecta As COMConecta.DCOMConecta

Dim lsSQL As String
    
    On Error GoTo eliminarCoberturasAgenciasError
    
    lsSQL = "exec stp_del_RFC0772012_EliminarCoberturasAgencias " & pnIdCobAgePer & ""
    Set oDCOMConecta = New COMConecta.DCOMConecta
    
    oDCOMConecta.AbreConexion
    oDCOMConecta.CargaRecordSet (lsSQL)
    oDCOMConecta.CierraConexion
        
    Set oDCOMConecta = Nothing
    Exit Sub
eliminarCoberturasAgenciasError:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Public Sub eliminarCoberturaAgenciaPeriodo(ByVal pnIdCobAgePer As Integer)
Dim oDCOMConecta As COMConecta.DCOMConecta

Dim lsSQL As String
    
    On Error GoTo eliminarCoberturaAgenciaPeriodoError
    
    lsSQL = "exec stp_del_RFC0772012_EliminarCoberturaAgenciaPeriodo " & pnIdCobAgePer & ""
    Set oDCOMConecta = New COMConecta.DCOMConecta
    
    oDCOMConecta.AbreConexion
    oDCOMConecta.CargaRecordSet (lsSQL)
    oDCOMConecta.CierraConexion
        
    Set oDCOMConecta = Nothing
    Exit Sub
eliminarCoberturaAgenciaPeriodoError:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Public Function verificarRangoCoberturaAgencia(ByVal pnMesIni As Integer, _
                                               ByVal pnAnioIni As Integer, _
                                               ByVal pnMesFin As Integer, _
                                               ByVal pnAnioFin As Integer, _
                                               Optional ByVal pbNuevo As Boolean = True, _
                                               Optional ByVal pnIdCobAgePer As Long = 0) As Boolean
Dim oDCOMConecta As COMConecta.DCOMConecta
Dim rsAgencias As ADODB.Recordset
Dim lsSQL As String
    
    On Error GoTo verificarRangoCoberturaAgenciaError
    
    lsSQL = "exec stp_sel_RFC0772012_VerificarRangoCoberturaAgencia '" & pnMesIni & "', '" & pnAnioIni & "', '" & pnMesFin & "', '" & pnAnioFin & "', " & IIf(pbNuevo, 1, 0) & ", " & pnIdCobAgePer & ""
    Set oDCOMConecta = New COMConecta.DCOMConecta
    oDCOMConecta.AbreConexion
    Set rsAgencias = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion
    
    
    If Not rsAgencias.EOF And Not rsAgencias.BOF Then
        If rsAgencias!nNroRango > 0 Then
            verificarRangoCoberturaAgencia = False
        Else
            verificarRangoCoberturaAgencia = True
        End If
    Else
    
        verificarRangoCoberturaAgencia = True
    End If

    Set rsAgencias = Nothing
    Set oDCOMConecta = Nothing
    Exit Function
verificarRangoCoberturaAgenciaError:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Public Function devolverCoberturasAgencias(ByVal pnIdCobAgePer As Integer) As ADODB.Recordset
Dim oDCOMConecta As COMConecta.DCOMConecta
Dim rsAgencias As ADODB.Recordset
Dim lsSQL As String
    
    On Error GoTo devolverCoberturasAgenciasError
    
    lsSQL = "exec stp_sel_RFC0772012_DevolverCoberturasAgencias " & pnIdCobAgePer
    Set oDCOMConecta = New COMConecta.DCOMConecta
    oDCOMConecta.AbreConexion
    Set rsAgencias = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion
    
    
    If Not rsAgencias.EOF And Not rsAgencias.BOF Then
        Set devolverCoberturasAgencias = rsAgencias
    End If

    Set rsAgencias = Nothing
    Set oDCOMConecta = Nothing
    Exit Function
devolverCoberturasAgenciasError:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Public Function devolverPeriodos() As ADODB.Recordset
Dim oDCOMConecta As COMConecta.DCOMConecta
Dim rsAgencias As ADODB.Recordset
Dim lsSQL As String
    
    On Error GoTo devolverPeriodosError
    
    lsSQL = "exec stp_sel_RFC0772012_DevolverPeriodos"
    Set oDCOMConecta = New COMConecta.DCOMConecta
    oDCOMConecta.AbreConexion
    Set rsAgencias = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion
    
    
    If Not rsAgencias.EOF And Not rsAgencias.BOF Then
        Set devolverPeriodos = rsAgencias
    End If

    Set rsAgencias = Nothing
    Set oDCOMConecta = Nothing
    Exit Function
devolverPeriodosError:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Public Function actualizarCoberturaAgencia(ByVal psAgeCod As String, _
                                           ByVal pnTipoCobertura As Integer, _
                                           ByVal pnIdCobAgePer As Integer, _
                                           ByVal pnMoneda As Integer, _
                                           ByVal pnCobertura As Currency, _
                                           ByVal psMovNro As String, _
                                           ByVal pnIdCoberturaAgencia As Integer) As Integer
Dim oDCOMConecta As COMConecta.DCOMConecta
Dim rsCobertura As ADODB.Recordset
Dim lsSQL As String
    
    On Error GoTo actualizarCoberturaAgenciaError
    
    lsSQL = "exec stp_upd_RFC0772012_ActualizarCoberturaAgencia '" & psAgeCod & "', " & pnTipoCobertura & ", " & pnIdCobAgePer & ", " & pnMoneda & ", " & pnCobertura & ", '" & psMovNro & "', " & pnIdCoberturaAgencia & ""
    Set oDCOMConecta = New COMConecta.DCOMConecta
    Set rsCobertura = New ADODB.Recordset
    
    oDCOMConecta.AbreConexion
    
    Set rsCobertura = oDCOMConecta.CargaRecordSet(lsSQL)
    
    If Not rsCobertura.BOF And Not rsCobertura.EOF Then
        actualizarCoberturaAgencia = rsCobertura!IdCoberturaAgencia
    Else
        actualizarCoberturaAgencia = 0
    End If
    
    oDCOMConecta.CierraConexion
        
    Set rsCobertura = Nothing
    Set oDCOMConecta = Nothing
    Exit Function
actualizarCoberturaAgenciaError:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Public Function actualizarCoberturaAgenciaPeriodo(ByVal pnMesDes As Integer, _
                                                  ByVal pnAnioDes As Integer, _
                                                  ByVal pnMesHas As Integer, _
                                                  ByVal pnAnioHas As Integer, _
                                                  ByVal psMovNro As String, _
                                                  ByVal pnIdCobAgePer As Integer) As Integer
Dim oDCOMConecta As COMConecta.DCOMConecta
Dim rsCobertura As ADODB.Recordset
Dim lsSQL As String
    
    On Error GoTo actualizarCoberturaAgenciaPeriodoError
    
    lsSQL = "exec stp_upd_RFC0772012_ActualizarCoberturaAgenciaPeriodo " & pnMesDes & ", " & pnAnioDes & ", " & pnMesHas & ", " & pnAnioHas & ",  '" & psMovNro & "', " & pnIdCobAgePer & ""
    Set oDCOMConecta = New COMConecta.DCOMConecta
    Set rsCobertura = New ADODB.Recordset
    
    oDCOMConecta.AbreConexion
    
    Set rsCobertura = oDCOMConecta.CargaRecordSet(lsSQL)
    
    If Not rsCobertura.BOF And Not rsCobertura.EOF Then
        actualizarCoberturaAgenciaPeriodo = rsCobertura!IdCobAgePer
    Else
        actualizarCoberturaAgenciaPeriodo = 0
    End If
    
    oDCOMConecta.CierraConexion
        
    Set rsCobertura = Nothing
    Set oDCOMConecta = Nothing
    Exit Function
actualizarCoberturaAgenciaPeriodoError:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
'***Fin Agregado por ELRO el 20120723*******************
'JUEZ 20131007 ********************************************************
Public Function GetUsuariosArbol(ByVal psAreaCod As String, ByVal psAgeCod As String) As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String

Set oCon = New COMConecta.DCOMConecta
If oCon.AbreConexion = False Then Exit Function

    sSql = "exec stp_sel_DevolverUsuariosArbol '" & psAreaCod & "','" & psAgeCod & "'"

Set GetUsuariosArbol = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function
'END JUEZ *************************************************************
'RECO20140621 ERS095-2014***************************************
Public Function DevuelveListaAgenciaJT(ByVal psUser As String) As ADODB.Recordset
    Dim rsVar As New ADODB.Recordset
    Dim sSql As String, sSqlaux As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function

    sSql = "stp_sel_DevuelveListaAgenciaJT " & psUser
    
    Set rsVar = oCon.CargaRecordSet(sSql)
    Set DevuelveListaAgenciaJT = rsVar
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'JATO 20210105 *****
Public Function GetUserUltimoArqueo(ByVal psFechaSist, ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal pnTipoAgencia As Integer) As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function
    Sql = " exec stp_sel_ObtenerUsuariosArqueo '" & Format(psFechaSist, "yyyyMMdd") & "', '" & psAreaCod & "','" & psAgeCod & "'," & pnTipoAgencia
Set rs = oCon.CargaRecordSet(Sql)
Set GetUserUltimoArqueo = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
'RECO FIN*******************************************************

'RIRO20140705 ERS072 *****
Public Function GetUserAreaAgenciaRelacion(ByVal psFechaSist As String, ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal pnTipoAgencia As Integer) As ADODB.Recordset
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function
Sql = " exec stp_sel_usuariosxarea '" & Format(psFechaSist, "yyyyMMdd") & "','" & psAreaCod & "','" & psAgeCod & "'," & pnTipoAgencia
Set rs = oCon.CargaRecordSet(Sql)
Set GetUserAreaAgenciaRelacion = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetUserAreaAgenciaRelacionXuser(ByVal psUser As String, ByVal psFechaSist As String, ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal pnTipoAgencia As Integer, ByVal pnTipoUsuario As Integer) As ADODB.Recordset
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function
Sql = " exec STP_SEL_USUARIO_ARQUEADO '" & psUser & "', '" & Format(psFechaSist, "yyyyMMdd") & "','" & psAreaCod & "','" & psAgeCod & "'," & pnTipoAgencia & "," & pnTipoUsuario

Set rs = oCon.CargaRecordSet(Sql)
Set GetUserAreaAgenciaRelacionXuser = rs
oCon.CierraConexion
Set oCon = Nothing
End Function

'TORE ERS033-2018
Public Function GetUserAreaAgenciaRelacionXuser2(ByVal psUser As String, ByVal psFechaSist As String, ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal pnTipoAgencia As Integer, ByVal pnTipoUsuario As Integer) As ADODB.Recordset
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function
Sql = " exec STP_SEL_USUARIO_ARQUEADO2 '" & psUser & "', '" & Format(psFechaSist, "yyyyMMdd") & "','" & psAreaCod & "','" & psAgeCod & "'," & pnTipoAgencia & "," & pnTipoUsuario
Set rs = oCon.CargaRecordSet(Sql)
Set GetUserAreaAgenciaRelacionXuser2 = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function ActualizaUsuarioArquedor(ByVal pnIdArqueo As Integer, ByVal psUsuario As String, ByVal psMovNro As String, ByVal pnTipoProceso As Integer) As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function
    Sql = "exec stp_upd_ActualizarUsuarioArqueador " & pnIdArqueo & ",'" & psUsuario & "','" & psMovNro & "'," & pnTipoProceso
    Set rs = oCon.CargaRecordSet(Sql)
    Set ActualizaUsuarioArquedor = rs
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'END TORE


Public Function SetRelacionArqueo(ByVal psRelacionUsuarios As String, ByVal psUserSinPar As String, ByVal pnEstado As Integer, ByVal psMovNroRegistro As String, _
                                  ByVal psUsuariosExcluidos As String) As Boolean

Dim Sql As String
Dim sUsuariosRelacion() As String
Dim sUsuario As String, sUsuarioArqueador As String
Dim i As Integer, nEstado As Integer
Dim bTrans As Boolean
Dim rs As New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta

On Error GoTo Error
   
    If oCon.AbreConexion = False Then Exit Function
    
    sUsuariosRelacion = Split(psRelacionUsuarios, ",")
    nEstado = 1 ' GENERADO
    oCon.BeginTrans
    bTrans = True

    'Insertando las parejas de arqueo
    For i = 0 To UBound(sUsuariosRelacion)
        sUsuario = Split(sUsuariosRelacion(i), "-")(0)
        sUsuarioArqueador = Split(sUsuariosRelacion(i), "-")(1)
        Sql = " EXEC STP_INS_ARQUEOENTREVENTANILLA '" & sUsuario & "','" & sUsuarioArqueador & "'," & pnEstado & ", '" & psMovNroRegistro & "'"
        oCon.CargaRecordSet (Sql)
        Sql = " EXEC STP_INS_ARQUEOENTREVENTANILLA '" & sUsuarioArqueador & "','" & sUsuario & "'," & pnEstado & ", '" & psMovNroRegistro & "'"
        oCon.CargaRecordSet (Sql)
    Next
    
    'Insertando el usuario impar, si hubiese
    If Len(Trim(psUserSinPar)) > 0 Then
        Sql = " EXEC STP_INS_ARQUEOENTREVENTANILLA '" & psUserSinPar & "',''," & pnEstado & ", '" & psMovNroRegistro & "'"
        oCon.CargaRecordSet (Sql)
    End If
    
    'insertando los usuarios excluidos
    If Len(Trim(psUsuariosExcluidos)) > 0 Then
        sUsuariosRelacion() = Split(psUsuariosExcluidos, ",")
        For i = 0 To UBound(sUsuariosRelacion)
            Sql = " EXEC STP_INS_ARQUEOENTREVENTANILLA '" & sUsuariosRelacion(i) & "','',6, '" & psMovNroRegistro & "'"
            oCon.CargaRecordSet (Sql)
        Next
    End If
    
    oCon.CommitTrans
    bTrans = False

    oCon.CierraConexion
    Set oCon = Nothing
    SetRelacionArqueo = True
    Exit Function
    
Error:
    If bTrans Then
        oCon.RollbackTrans
        Set oCon = Nothing
    End If
    SetRelacionArqueo = False

End Function
'END RIRO ****************
'RECO20141011 ERS022-2014***************************************************
Public Function RegistraLimiteEfectivo(ByVal pnGrupo As Integer, ByVal pnAnioDesde As Integer, ByVal pnMesDesde As Integer, ByVal pnAnioHasta As Integer, ByVal pnMesHasta As Integer) As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    On Error GoTo ErrorRegistraLimiteEfectivo
    If oCon.AbreConexion = False Then Exit Function
    Sql = "exec stp_ins_RegistraLimiteEfectivo " & pnGrupo & "," & pnAnioDesde & "," & pnMesDesde & "," & pnAnioHasta & "," & pnMesHasta
    Set rs = oCon.CargaRecordSet(Sql)
    Set RegistraLimiteEfectivo = rs
    oCon.CierraConexion
    Exit Function
ErrorRegistraLimiteEfectivo:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
Public Sub RegistraLimiteEfectivoDet(ByVal pnCodigoID As Integer, ByVal psAgeCod As String, ByVal pnPoliza As Double, ByVal pnMaximo As Double, ByVal pnMinimo As Double)
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
        
        On Error GoTo ErrorRegistraLimiteEfectivoDet
        
        lsSQL = "exec stp_ins_RegistraLimiteEfectivoDet " & pnCodigoID & ",'" & psAgeCod & "'," & pnPoliza & "," & pnMaximo & "," & pnMinimo
        Set oDCOMConecta = New COMConecta.DCOMConecta
        
        oDCOMConecta.AbreConexion
        oDCOMConecta.CargaRecordSet (lsSQL)
        oDCOMConecta.CierraConexion
            
        Set oDCOMConecta = Nothing
        Exit Sub
ErrorRegistraLimiteEfectivoDet:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Public Function ObtieneHistorialLimiteEfectivo() As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    On Error GoTo ErrorObtieneHistorialLimiteefectivo
    If oCon.AbreConexion = False Then Exit Function
    Sql = "exec stp_sel_ObtieneHistorialLimiteefectivo "
    Set rs = oCon.CargaRecordSet(Sql)
    Set ObtieneHistorialLimiteEfectivo = rs
    oCon.CierraConexion
    Exit Function
ErrorObtieneHistorialLimiteefectivo:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
Public Function ObtieneDatosLimiteEfectivo(ByVal pnCodigoID) As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    On Error GoTo ErrorObtieneDatosLimiteEfectivo
    If oCon.AbreConexion = False Then Exit Function
    Sql = "exec stp_sel_ObtieneDatosLimiteEfectivo " & pnCodigoID
    Set rs = oCon.CargaRecordSet(Sql)
    Set ObtieneDatosLimiteEfectivo = rs
    oCon.CierraConexion
    Exit Function
ErrorObtieneDatosLimiteEfectivo:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
Public Sub EliminaLimiteEfectivo(ByVal pnCodigoID As Integer)
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
        
        On Error GoTo ErrorEliminaLimiteEfectivoDet
        
        lsSQL = "exec stp_upd_EliminaLimiteEfectivoDet " & pnCodigoID
        Set oDCOMConecta = New COMConecta.DCOMConecta
        
        oDCOMConecta.AbreConexion
        oDCOMConecta.CargaRecordSet (lsSQL)
        oDCOMConecta.CierraConexion
            
        Set oDCOMConecta = Nothing
        Exit Sub
ErrorEliminaLimiteEfectivoDet:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
'RECO END******************************************************************
'FRHU 20150326
Public Function GetResponsableComite() As ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Dim sSql As String
    
    Set oCon = New COMConecta.DCOMConecta
    If oCon.AbreConexion = False Then Exit Function
    sSql = "exec stp_sel_ResponsableComite"
    Set GetResponsableComite = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'FIN FRHU 20150326

'JIPR20190109 ERS077-2018
Public Function GetTarifario(ByVal nConstante As ConstanteCabecera) As ADODB.Recordset
    Dim sSql As String
    Dim rsRel As New ADODB.Recordset
    rsRel.CursorLocation = adUseClient
    
    sSql = "exec stp_sel_CapConfigTarifario " & nConstante
        
    rsRel.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsRel.ActiveConnection = Nothing
    Set GetTarifario = rsRel
    Set rsRel = Nothing
End Function


