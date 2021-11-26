VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMCaptaServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public dbCmact As Connection
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String
Dim sSQL As String

Dim oFun As New COMFunciones.FCOMVarPublicas

Public Sub GrabaMovCobranzas(ByVal nMovNro As Long, ByVal cPersCod As String, ByVal cCodigo As String, ByVal nTpoCuenta As Integer, ByVal nMonto As Double, ByVal cReferencia As String)

sSQL = "Insert movconvcobranza(nMovNro ,cPersCod ,cCodigo ,nTpoCuenta ,nMonto ,cReferencia)  " _
       & " VALUES (" & nMovNro & ",'" & cPersCod & "','" & cCodigo & "'," & nTpoCuenta & "," & nMonto & ",'" & cReferencia & "' ) "
       
dbCmact.Execute sSQL
    
    
End Sub

Public Function GetnMovNro(ByVal sMovNro As String) As Long
Dim rsMov As ADODB.Recordset
sSQL = "Select nMovNro From Mov Where cMovNro = '" & sMovNro & "'"
Set rsMov = New ADODB.Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
If rsMov.EOF And rsMov.BOF Then
    GetnMovNro = 0
Else
    GetnMovNro = rsMov("nMovNro")
End If
rsMov.Close
Set rsMov = Nothing
End Function

Public Function GetServicioParametros(ByVal nServicio As CaptacInstServicios) As ADODB.Recordset
Dim rsServ As ADODB.Recordset
Set rsServ = New ADODB.Recordset
rsServ.CursorLocation = adUseClient
sSQL = "Select nComision, cCtaCodAbono, nTipoComision From " _
    & "CapServicioParam Where nServCod = " & nServicio
rsServ.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsServ.ActiveConnection = Nothing
Set GetServicioParametros = rsServ
Set rsServ = Nothing
End Function

Public Sub ActualizaServicioParametro(ByVal nServicio As CaptacInstServicios, _
        ByVal sCuenta As String, ByVal nComision As Double, ByVal nTipoComision As CapServTipoComision)
sSQL = "Update CapServicioParam Set nComision = " & nComision & ", " _
    & "cCtaCodAbono = '" & sCuenta & "', " _
    & "nTipoComision = " & nTipoComision & " " _
    & "Where nServCod = " & nServicio
dbCmact.Execute sSQL
End Sub

Public Sub AgregaMovLavDinero(ByVal nMovNro As Long, ByVal sPersCod As String)
sSQL = "INSERT MovLavDinero (nMovNro,cPersCod) " _
    & "VALUES (" & nMovNro & ",'" & sPersCod & "')"
dbCmact.Execute sSQL
End Sub

Public Sub AgregaServicioParametro(ByVal nServicio As CaptacInstServicios, _
        ByVal sCuenta As String, ByVal nComision As Double, ByVal nTipoComision As CapServTipoComision)
sSQL = "Insert CapServicioParam (nServCod,nComision,cCtaCodAbono,nTipoComision)  " _
    & "VALUES (" & nServicio & "," & nComision & ",'" & sCuenta & "'," & nTipoComision & ")"
dbCmact.Execute sSQL
End Sub

Public Function EsReciboCobrado(ByVal nOperacion As CaptacOperacion, ByVal sNumRecibo As String) As Boolean
Dim rsRec As ADODB.Recordset
Set rsRec = New ADODB.Recordset
rsRec.CursorLocation = adUseClient
sSQL = "Select M.cMovNro, S.cNumRecibo, S.cCodigo From Mov M INNER JOIN MovServicios S " _
    & "ON M.nMovNro = S.nMovNro Where S.cNumRecibo = '" & sNumRecibo & "' And " _
    & "M.cOpeCod = '" & nOperacion & "'"
rsRec.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsRec.ActiveConnection = Nothing
If rsRec.EOF And rsRec.BOF Then
    EsReciboCobrado = False
Else
    EsReciboCobrado = True
End If
rsRec.Close
Set rsRec = Nothing
End Function

Public Sub ActualizaEstadoMov(ByVal nMovNro As Long, ByVal nMovEstado As MovEstado, ByVal nMovFlag As MovFlag)
sSQL = "UPDATE Mov Set nMovEstado = " & nMovEstado & ", nMovFlag = " & nMovFlag & " " _
    & "WHERE nMovNro = '" & nMovNro & "'"
dbCmact.Execute sSQL
End Sub

Public Sub AgregaMov(ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
        ByVal sGlosa As String, Optional nMovEstado As MovEstado = gMovEstContabMovContable, _
        Optional nMovFlag As MovFlag = gMovFlagVigente)
sSQL = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & nOperacion & "','" & sGlosa & "'," & nMovEstado & "," & nMovFlag & ")"
dbCmact.Execute sSQL
End Sub

Public Sub AgregaMovRef(ByVal nMovNro As Long, ByVal nMovNroRef As Long)
sSQL = "INSERT MovRef (nMovNro,nMovNroRef) " _
    & "VALUES (" & nMovNro & "," & nMovNroRef & ")"
dbCmact.Execute sSQL
End Sub
Public Sub AgregaMovServicios(ByVal nMovNro As Long, ByVal sNumRecibo As String, _
        ByVal sCodigo As String, ByVal nMonto As Double, _
        Optional nFlag As CaptacFlagServicios = gCapServFlagRegistrado, _
        Optional nmoneda As Moneda = gMonedaNacional, Optional gcOpeCod As String = "")
sSQL = "INSERT MovServicios (nMovNro,cNumRecibo,cCodigo,nMonto,nFlag,nMoneda, cOpeCod) " _
    & "VALUES (" & nMovNro & ",'" & sNumRecibo & "','" & sCodigo & "'," & nMonto & "," & nFlag & "," & nmoneda & ", '" & gcOpeCod & "')"
dbCmact.Execute sSQL
End Sub

Public Sub AgregaMovConvCobranza(ByVal nMovNro As Long, ByVal cPersCod As String, ByVal cCodigo As String, ByVal nTpoCuenta As Integer, ByVal nMonto As Double, ByVal cReferencia As String)
sSQL = "INSERT movconvcobranza (nMovNro,cPersCod,cCodigo,nTpoCuenta,nMonto,cReferencia) " _
    & "VALUES (" & nMovNro & ",'" & cPersCod & "','" & cCodigo & "'," & nTpoCuenta & "," & nMonto & ",'" & cReferencia & "')"
dbCmact.Execute sSQL
End Sub

Public Sub AgregaMovServiciosDet(ByVal nMovNro As Long, ByVal sNumRecibo As String, _
        ByVal nConcepto As CaptacConcepto, ByVal nMonto As Double, Optional nCuota As Long = 0, Optional gcOpeCod As String = "")

sSQL = "INSERT MovServiciosDet (nMovNro,cNumRecibo,nPrdConcepto,nNroCuota,nMonto, cOpeCod) " _
    & "VALUES (" & nMovNro & ",'" & sNumRecibo & "'," & nConcepto & "," & nCuota & "," & nMonto & ", '" & gcOpeCod & "')"
dbCmact.Execute sSQL
End Sub

Public Function ActualizaFlagServicios(ByVal nMovNro As Long, ByVal nFlag As CaptacFlagServicios)
sSQL = "UPDATE MovServicios Set nFlag  = " & nFlag & " WHERE nMovNro = " & nMovNro
dbCmact.Execute sSQL
End Function

Public Function GetMovServicios(ByVal nServicio As CaptacOperacion, _
        ByVal dInicio As Date, ByVal dFin As Date, Optional bReciboYaAbono As Boolean = False) As ADODB.Recordset
Dim rsMov As ADODB.Recordset
Dim sRef As String
Set rsMov = New ADODB.Recordset

sRef = ""
If Not bReciboYaAbono Then
    sRef = "And M.nMovNro NOT IN (Select MR.nMovNroRef From MovRef MR JOIN Mov M1 ON " _
        & "MR.nMovNro = M1.nMovNro Where M.nMovFlag IN (" & gMovFlagVigente & "))"
End If

sSQL = "Select dFecha = CONVERT(Varchar(10),CONVERT(Datetime,LEFT(M.cMovNro,8),103)) + ' ' + " _
    & "SUBSTRING(M.cMovNro,9,2) + ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2), " _
    & "S.cNumRecibo, S.cCodigo, SD.nMonto, A.cAgeDescripcion, RIGHT(M.cMovNro,4) Usuario, M.nMovNro " _
    & "From Mov M INNER JOIN MovServicios S JOIN MovServiciosDet SD ON S.nMovNro = SD.nMovNro And " _
    & "S.cNumRecibo = SD.cNumRecibo ON M.nMovNro = S.nMovNro JOIN Agencias A ON " _
    & "SUBSTRING(M.cMovNro,18,2) = A.cAgeCod Where LEFT(M.cMovNro,8) BETWEEN '" _
    & Format$(dInicio, "yyyymmdd") & "' And '" & Format$(dFin, "yyyymmdd") & "' And M.nMovFlag IN (" & gMovFlagVigente & ") " _
    & "And M.cOpeCod = '" & nServicio & "' And SD.nPrdConcepto = " & gConcCapital & " " & sRef
    
rsMov.CursorLocation = adUseClient
rsMov.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetMovServicios = rsMov
Set rsMov = Nothing
End Function

Public Sub CreaTablaDBFSedalib(ByVal sNombreArchivo As String)
Dim dbAux As Connection
Set dbAux = New Connection

sSQL = "CREATE TABLE " & sNombreArchivo & " ( "
sSQL = sSQL & "FACSERNRO N(2,0) NOT NULL, "
sSQL = sSQL & "FACNRO N(7,0)  NOT NULL, "
sSQL = sSQL & "HISEMIFEC D(8) NOT NULL, "
sSQL = sSQL & "HISCOBFEC D(8) NOT NULL, "
sSQL = sSQL & "HISCORNRO N(5,0) NOT NULL, "
sSQL = sSQL & "OFICOD N(1,0) NOT NULL, "
sSQL = sSQL & "OFIAGECOD N(0,0) NOT NULL, "
sSQL = sSQL & "HISFACANO C(4) NOT NULL, "
sSQL = sSQL & "HISFACMES C(2) NOT NULL, "
sSQL = sSQL & "CLIUNICOD C(11) NOT NULL, "
sSQL = sSQL & "HISCHEDIG N(1,0) NOT NULL, "
sSQL = sSQL & "HISDIGCHE N(1,0) NOT NULL, "
sSQL = sSQL & "HISCOBHRS C(8) NOT NULL, "
sSQL = sSQL & "HISFLAGZ N(0,0) NOT NULL, "
sSQL = sSQL & "HISACTFEC D(8) NOT NULL, "
sSQL = sSQL & "HISACTFLA N(0,0) NOT NULL, "
sSQL = sSQL & "HISCTRFLA N(0,0) NOT NULL, "
sSQL = sSQL & "HISPAGO N(8,2) NOT NULL, "
sSQL = sSQL & "HISCOBTIP  C(1) NOT NULL, "
sSQL = sSQL & "HISEXCCOB N(10,2) NOT NULL,"
sSQL = sSQL & "HISCODZON N(0,0) NOT NULL, "
sSQL = sSQL & "HISCODREG N(0,0) NOT NULL, "
sSQL = sSQL & "HISDIGCOD N(1,0) NOT NULL, "
sSQL = sSQL & "HISCOBDIA N(11,2) NOT NULL, "
sSQL = sSQL & "HISGRUCOD N(2,0) NOT NULL, "
sSQL = sSQL & "HISGRUSUB N(3,0) NOT NULL, "
sSQL = sSQL & "NFILE  C(2) NOT NULL )"
    
dbAux.Open oFun.gsConnServDBF
dbAux.Execute sSQL
dbAux.Close
Set dbAux = Nothing
End Sub

Public Sub AgregaServDBF(ByVal sArchivo As String, ByVal sRecibo As String, _
        ByVal sCodigo As String, ByVal nMonto As Double, ByVal dFecha As Date, _
        ByVal sHora As String)
Dim dbAux As Connection
Set dbAux = New Connection
sSQL = "INSERT INTO " & sArchivo & " (Facsernro,Facnro,hisemifec,Hiscobfec,hiscornro,oficod,ofiagecod,hisfacano,hisfacmes,cliunicod,hischedig,hisdigche,hiscobhrs,hisflagz,hisactfec,hisactfla,hisctrfla,hispago,hiscobtip,hisexccob,hiscodzon,hiscodreg,hisdigcod,hiscobdia,hisgrucod,hisgrusub,nfile) " _
    & "VALUES(" & Left(sRecibo, 3) & "," & Mid(sRecibo, 4, 8) & ",CTOD('" & Format(dFecha, "mm/dd/yyyy") & "'),CTOD('" & Format(dFecha, "mm/dd/yyyy") & "'),0,5,0,'','','" & sCodigo & "',0,0,'" & sHora & "',0," _
    & "CTOD('" & Format(dFecha, "mm/dd/yyyy") & "'),0,0," & Format(nMonto, "##0.00") & ",'',0,0,0,4,0,0,0,'')"
dbAux.Open oFun.gsConnServDBF
dbAux.Execute sSQL
dbAux.Close
Set dbAux = Nothing
End Sub
Public Function GetServConvenios2() As ADODB.Recordset
Dim rsConv As ADODB.Recordset
Set rsConv = New ADODB.Recordset
rsConv.CursorLocation = adUseClient
sSQL = "Select distinct Codigo=C.cPersCod, Institucion=rtrim(ltrim(P.cPersNombre)) , Tipo=case when c.nconvcod=101 then 'OTROS INGRESOS CON PERSONAS' " _
    & " else 'COLEGIOS' end, c.nconvcod " _
    & "  From CaptacConvenio C INNER JOIN " & sDBPersona & "Persona P " _
    & "ON C.cPersCod = P.cPersCod INNER JOIN " & sDBComunes & "Constante K ON C.nConvCod = K.nConsValor " _
    & "And K.nConscod =2015  order by rtrim(ltrim(p.cpersnombre)) "
rsConv.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsConv.ActiveConnection = Nothing
Set GetServConvenios2 = rsConv
Set rsConv = Nothing
End Function

Public Function GetServConvenios(Optional ByVal nTipoCOnv As Integer) As ADODB.Recordset
Dim rsConv As ADODB.Recordset
Set rsConv = New ADODB.Recordset
rsConv.CursorLocation = adUseClient
sSQL = "Select distinct Codigo=C.cPersCod, Institucion=ltrim(rtrim(P.cPersNombre)), cpersid=isnull(( Select top 1 cpersidnro from persid g where cpersidtpo in ('1','2') and g.cperscod=p.cperscod ),'')" _
    & "  From CaptacConvenio C INNER JOIN " & sDBPersona & "Persona P " _
    & "ON C.cPersCod = P.cPersCod INNER JOIN " & sDBComunes & "Constante K ON C.nConvCod = K.nConsValor " _
    & "And K.nConsValor = " & nTipoCOnv & "" _
    & " order by ltrim(rtrim(P.cPersNombre))  "
    
rsConv.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsConv.ActiveConnection = Nothing
Set GetServConvenios = rsConv
Set rsConv = Nothing
End Function

Public Function GetServConveniosArbol() As ADODB.Recordset
Dim rsConv As ADODB.Recordset
Set rsConv = New ADODB.Recordset
rsConv.CursorLocation = adUseClient
sSQL = "Select C.cPersCod, P.cPersNombre, nNivel = 1 " _
    & "From CaptacConvenio C INNER JOIN " & sDBPersona & "Persona P " _
    & "ON C.cPersCod = P.cPersCod "
rsConv.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsConv.ActiveConnection = Nothing
Set GetServConveniosArbol = rsConv
Set rsConv = Nothing
End Function

Public Function GetServConveniosCodArbol() As ADODB.Recordset
Dim rsConv As ADODB.Recordset
Set rsConv = New ADODB.Recordset
rsConv.CursorLocation = adUseClient
sSQL = "Select C.nConvCod, P.cPersNombre, nNivel = 1 " _
    & "From CaptacConvenio C INNER JOIN " & sDBPersona & "Persona P " _
    & "ON C.cPersCod = P.cPersCod "
rsConv.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsConv.ActiveConnection = Nothing
Set GetServConveniosCodArbol = rsConv
Set rsConv = Nothing
End Function

Public Sub EliminaServConvenio(ByVal sPersona As String)
sSQL = "Delete CaptacConvenio Where cPersCod = '" & sPersona & "'"
dbCmact.Execute sSQL
End Sub

Public Sub EliminaServConvCuentas(ByVal sPersona As String)
sSQL = "Delete CaptacConvenioCta Where cPersCod = '" & sPersona & "'"
dbCmact.Execute sSQL
End Sub
Public Sub ActualizaServConvenio(ByVal sPersona As String, ByVal nConvenio As CaptacConvenios)
sSQL = " Update CaptacConvenio  set nConvCod=" & nConvenio _
    & " where cperscod='" & sPersona & "'"
dbCmact.Execute sSQL
End Sub
Public Sub AgregaServConvenio(ByVal sPersona As String, ByVal nConvenio As CaptacConvenios)
sSQL = "Insert CaptacConvenio (cPersCod,nConvCod) " _
    & "VALUES ('" & sPersona & "'," & nConvenio & ")"
dbCmact.Execute sSQL
End Sub

Public Sub AgregaServConvenioCuenta(ByVal sPersona As String, ByVal sCuenta As String, _
            ByVal nTipoCuenta As CaptacConvTipoCuenta, nMontoMoraDia As Currency)
sSQL = "Insert CaptacConvenioCta (cPersCod,cCtaCod,nTpoCuenta,nMontoMoraDia) " _
    & "VALUES ('" & sPersona & "','" & sCuenta & "'," & nTipoCuenta & "," & nMontoMoraDia & " )"
dbCmact.Execute sSQL
End Sub

Public Function GetServPlanPagos(Optional sPersona As String = "", Optional nConvenio As CaptacConvenios) As ADODB.Recordset
Dim rsPlan As ADODB.Recordset

If sPersona <> "" Then
    sSQL = "Select CONVERT(Varchar(10),dFecVenc,103) cVencimiento, nMontoCuota, nMontoGasto, " _
        & "bMora = CASE WHEN bAfectoMora = 1 THEN 'SI' ELSE 'NO' END, bFeriado = CASE WHEN bAfectoFeriado = 1 THEN 'SI' ELSE 'NO' END " _
        & "From CaptacConvenioPlanPag WHERE cPersCod = '" & sPersona & "'"
Else
    sSQL = "Select PP.dFecVenc, PP.nMontoCuota, PP.nMontoGasto, PP.bAfectoMora, PP.bAfectoFeriado, PP.nNroCuota " _
        & "From CaptacConvenio C INNER JOIN CaptacConvenioPlanPag PP ON C.cPersCod = PP.cPersCod " _
        & "WHERE C.nConvCod = " & nConvenio & " ORDER BY PP.dFecVenc"
End If
Set rsPlan = New ADODB.Recordset
rsPlan.CursorLocation = adUseClient
rsPlan.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsPlan.ActiveConnection = Nothing
Set GetServPlanPagos = rsPlan
Set rsPlan = Nothing
End Function

Public Function GetServConvCuentas(Optional sPersona As String = "", Optional nConvenio As CaptacConvenios, Optional cSoloActivas As String = "N") As ADODB.Recordset
Dim rsPlan As ADODB.Recordset

If sPersona <> "" Then
  If cSoloActivas = "N" Then
    sSQL = "Select [Nro Cuenta]=C.cCtaCod,Moneda=case when substring(c.cctacod,9,1)='1' then 'SOLES' else 'DOLARES' end ,[Fecha Apertura]=convert(char(10),ca.dapertura,103)  " _
        & "FROM CaptacConvenioCta C INNER JOIN " & sDBComunes & "Constante K ON C.nTpoCuenta = K.nConsValor " _
        & " JOIN CAPTACIONES CA ON CA.CCTACOD=C.CCTACOD " _
        & "WHERE    C.cPersCod = '" & sPersona & "' And K.nConsCod = " & gCaptacConvTipoCuenta & " ORDER BY " _
        & "C.nTpoCuenta"
  Else
    sSQL = "Select [Nro Cuenta]=C.cCtaCod,Moneda=case when substring(c.cctacod,9,1)='1' then 'SOLES' else 'DOLARES' end ,[Fecha Apertura]=convert(char(10),ca.dapertura,103)  " _
        & "FROM CaptacConvenioCta C INNER JOIN " & sDBComunes & "Constante K ON C.nTpoCuenta = K.nConsValor " _
        & " JOIN CAPTACIONES CA ON CA.CCTACOD=C.CCTACOD " _
        & " join producto p on p.cctacod=ca.cctacod " _
        & "WHERE p.nprdestado not in (1300,1400) and C.cPersCod = '" & sPersona & "' And K.nConsCod = " & gCaptacConvTipoCuenta & " ORDER BY " _
        & "C.nTpoCuenta"
        
  End If
Else
    sSQL = "Select CC.cCtaCod, K.cConsDescripcion, CC.nTpoCuenta, CC.nMontoMoraDia " _
        & "FROM CaptacConvenioCta CC INNER JOIN CaptacConvenio C ON CC.cPersCod = C.cPersCod INNER JOIN " _
        & sDBComunes & "Constante K ON CC.nTpoCuenta = K.nConsValor WHERE C.nConvCod = " & nConvenio & " And " _
        & "K.nConsCod = " & gCaptacConvTipoCuenta & " ORDER BY CC.nTpoCuenta"
End If
Set rsPlan = New ADODB.Recordset
rsPlan.CursorLocation = adUseClient
rsPlan.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsPlan.ActiveConnection = Nothing
Set GetServConvCuentas = rsPlan
Set rsPlan = Nothing
End Function

Public Sub EliminaPlanPagos(ByVal sPersona As String)
sSQL = "Delete CaptacConvenioPlanPag Where cPersCod = '" & sPersona & "'"
dbCmact.Execute sSQL
End Sub

Public Sub AgregaPlanPagos(ByVal sPersona As String, ByVal nNroCuota As Long, _
        ByVal dVencimiento As Date, ByVal nMontoCuota As Double, ByVal nMontoGasto As Double, _
        ByVal bMora As Boolean, ByVal bFeriado As Boolean)
sSQL = "INSERT CaptacConvenioPlanPag (cPersCod,nNroCuota,dFecVenc,nMontoCuota,nMontoGasto,bAfectoMora,bAfectoFeriado) " _
    & "VALUES ( '" & sPersona & "'," & nNroCuota & ",'" & Format$(dVencimiento, "mm/dd/yyyy") & "'," & nMontoCuota & "," & nMontoGasto & "," & IIf(bMora, 1, 0) & "," & IIf(bFeriado, 1, 0) & ")"
dbCmact.Execute sSQL
End Sub

Private Sub Class_Initialize()
Dim sConn As String
Dim ClsIni As COMConecta.DCOMClasIni
Set ClsIni = New COMConecta.DCOMClasIni
sConn = ClsIni.CadenaConexion
sDBComunes = ClsIni.BaseComunes
sDBPersona = ClsIni.BasePersonas
sDBImagenes = ClsIni.BaseImagenes
Set ClsIni = Nothing
Set dbCmact = New Connection
dbCmact.Open sConn
dbCmact.Execute "Set DateFormat MDY"
End Sub

Private Sub Class_Terminate()
dbCmact.Close
Set dbCmact = Nothing
End Sub

Public Function GetUNTConceptosAlumno(ByVal sCurriculo As String, ByVal sTipoMatricula As String, _
        ByVal sIngresante As String, ByVal s2daProf As String) As ADODB.Recordset
Dim rsMat As ADODB.Recordset
sSQL = "Select C.nConceptoCod, C.cDescripcion Concepto, Monto = CASE CC.bPredefinido WHEN 1 THEN CC.nMonto WHEN 0 THEN 0 END FROM " _
    & "CapUNTConcepto C INNER JOIN CapUNTCurriculoConcepto CC INNER JOIN CapUNTMatricula M ON CC.cMatriculaCod = M.cMatriculaCod ON " _
    & "C.nConceptoCod = CC.nConceptoCod WHERE CC.cCurriculoTpo = '" & sCurriculo & "' " _
    & "AND CC.cMatriculaCod = '" & sTipoMatricula & "' And C.cEstado = '1' And C.nConceptoCod IN (" & gUNTMatReg & "," & gUNTMatExt & "," _
    & gUNTMatReg2daProf & "," & gUNTMatExt2daProf & ") And cSegProfesion = '" & s2daProf & "' And cIngresante = '" & sIngresante & "'"
Set rsMat = New ADODB.Recordset
rsMat.CursorLocation = adUseClient

rsMat.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMat.ActiveConnection = Nothing
Set GetUNTConceptosAlumno = rsMat
Set rsMat = Nothing
End Function

Public Function GetUNTAlumno(ByVal sMatricula As String) As ADODB.Recordset
Dim rsAlu As ADODB.Recordset
sSQL = "Select A.cNombre, E.cDescripcion, E.cEscuelaCod, E.cCurriculo from CapUNTAlumno A " _
    & "INNER JOIN CapUNTEscuela E ON A.cEscuelaCod = E.cEscuelaCod WHERE A.cMatricula = '" _
    & sMatricula & "'"
Set rsAlu = New ADODB.Recordset
rsAlu.CursorLocation = adUseClient
rsAlu.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsAlu.ActiveConnection = Nothing
Set GetUNTAlumno = rsAlu
Set rsAlu = Nothing
End Function

Public Function GetConvenioCuenta(ByVal nConvenio As CaptacConvenios) As ADODB.Recordset
Dim rsCta As ADODB.Recordset
sSQL = "Select CC.cCtaCod, K.cConsDescripcion, CC.nTpoCuenta, CC.nMontoMoraDia From " _
    & "CaptacConvenioCta CC INNER JOIN CaptacConvenio C ON CC.cPersCod = C.cPersCod INNER JOIN " _
    & "Where C.nConvCod = " & nConvenio
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetConvenioCuenta = rsCta
Set rsCta = Nothing
End Function

Public Function GetConvenioCuenta2(ByVal scperscod As String, ByVal nConvenio As CaptacConvenios) As ADODB.Recordset
Dim rsCta As ADODB.Recordset
sSQL = " Select CC.cCtaCod,cast( CC.nTpoCuenta as char(1)) + K.cConsDescripcion , CC.nMontoMoraDia, ESTADO=(case when p.nprdestado IN (1400,1300) then 'NO' ELSE 'SI' END) From " _
    & " CaptacConvenioCta CC INNER JOIN CaptacConvenio C ON CC.cPersCod = C.cPersCod INNER JOIN " _
    & " (select * from constante where nconscod=2016) K ON K.nconsvalor=cc.ntpocuenta  " _
    & " join producto p on p.cctacod=cc.cctacod " _
    & " Where C.nConvCod = " & nConvenio & " and cc.cperscod='" & scperscod & "'"
    
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetConvenioCuenta2 = rsCta
Set rsCta = Nothing
End Function


Public Function GetUNTCurriculoConcepto(ByVal sCurriculo As String, ByVal sTipoMatricula As String, _
        ByVal sIngresante As String, ByVal s2daProf As String) As ADODB.Recordset
Dim rsCurr As ADODB.Recordset
sSQL = "Select C.nConceptoCod, C.cDescripcion Concepto, Monto = CASE CC.bPredefinido WHEN 1 THEN CC.nMonto WHEN 0 THEN 0 END FROM " _
    & "CapUNTConcepto C INNER JOIN CapUNTCurriculoConcepto CC INNER JOIN CapUNTMatricula M ON CC.cMatriculaCod = M.cMatriculaCod ON " _
    & "C.nConceptoCod = CC.nConceptoCod WHERE CC.cCurriculoTpo = '" & sCurriculo & "' " _
    & "And CC.cMatriculaCod = '" & sTipoMatricula & "' And C.cEstado = '1' " _
    & "And cSegProfesion = '" & s2daProf & "' And cIngresante = '" & sIngresante & "'"
Set rsCurr = New ADODB.Recordset
rsCurr.CursorLocation = adUseClient
rsCurr.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCurr.ActiveConnection = Nothing
Set GetUNTCurriculoConcepto = rsCurr
Set rsCurr = Nothing
End Function

Public Function GetUNTConceptosOtros() As ADODB.Recordset
Dim rsCurr As ADODB.Recordset
sSQL = "Select C.nConceptoCod, C.cDescripcion Concepto, Monto = 0 FROM " _
    & "CapUNTConcepto C WHERE C.cEstado = '2'"
Set rsCurr = New ADODB.Recordset
rsCurr.CursorLocation = adUseClient
rsCurr.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCurr.ActiveConnection = Nothing
Set GetUNTConceptosOtros = rsCurr
Set rsCurr = Nothing
End Function

Public Function GetUNTAlumnoNombre(ByVal sAlumno As String) As ADODB.Recordset
Dim rsUNT As ADODB.Recordset
sSQL = "SELECT A.cMatricula [Num Matr], A.cNombre Nombre, E.cDescripcion Escuela, E.cEscuelaCod, E.cCurriculo from CapUNTAlumno A " _
    & "INNER JOIN CapUNTEscuela E ON A.cEscuelaCod = E.cEscuelaCod WHERE A.cNombre LIKE (RTRIM('" & sAlumno & "') + '%') " _
    & "ORDER BY A.cNombre"
Set rsUNT = New ADODB.Recordset
rsUNT.CursorLocation = adUseClient
rsUNT.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsUNT.ActiveConnection = Nothing
Set GetUNTAlumnoNombre = rsUNT
Set rsUNT = Nothing
End Function

Public Sub AgregaUNTOrdenPago(ByVal nMovNro As Long, ByVal sOrden As String, ByVal sTipo As String, _
        ByVal dFecPago As Date, ByVal sAlumno As String, ByVal nTotal As Double, _
        ByVal sAgencia As String)
sSQL = "INSERT CapUNTOrdenPago (nMovNro,cOrdenCod,cTipo,dFecPago,cMatricula,nMonto,nSaldo,cFlag,cAgencia) " _
    & "VALUES (" & nMovNro & ",'" & sOrden & "','" & sTipo & "','" & Format$(dFecPago, "mm/dd/yyyy") & "','" & sAlumno & "'," & Trim(nTotal) & ",0,'1','" & sAgencia & "')"
dbCmact.Execute sSQL
End Sub

Public Sub AgregaUNTOrdenPagoDet(ByVal nMovNro As Long, ByVal nConcepto As CapServUNTConcepto, _
        ByVal nMonto As Double)
sSQL = "INSERT CapUNTOrdenPagoDet (nMovNro,nConceptoCod,nMonto) " _
    & "VALUES (" & nMovNro & "," & nConcepto & "," & nMonto & ")"
dbCmact.Execute sSQL
End Sub
                
Public Function GetConvenioAlumno(ByVal nConvenio As CaptacConvenios, ByVal sCodigo As String) As ADODB.Recordset
Dim rsServ As ADODB.Recordset
Set rsServ = New ADODB.Recordset
rsServ.CursorLocation = adUseClient
sSQL = "Select CD.cCodigo, CD.cNombre, CD.cReferencia From CaptacConvenio C INNER JOIN CaptacConvenioDet CD " _
    & "ON C.cPersCod = CD.cPersCod WHERE C.nConvCod = " & nConvenio & " AND CD.cCodigo = '" & sCodigo & "'"
rsServ.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsServ.ActiveConnection = Nothing
Set GetConvenioAlumno = rsServ
Set rsServ = Nothing
End Function

Public Function GetGiroTarifario() As ADODB.Recordset
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
sSQL = "Select nMoneda, nMontoIni, nMontoFin, nValor, nTipoTarifa from GiroTarifario Order by nMoneda, nMontoIni"
rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetGiroTarifario = rs
Set rs = Nothing
End Function

Public Function GetGiroDatos(ByVal sCuenta As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
sSQL = "Select PD.cCtaCod, PD.dPrdEstado, LTRIM(RTRIM(K.cConsDescripcion)) cTipo, PD.nSaldo, P.cPersNombre cRemitente, P.cPersDireccDomicilio, " _
    & "D.cDestinatario, LTRIM(RTRIM(A.cAgeDescripcion)) cAgencia, ISNULL(ID.cPersIDnro,'') cDocID, D.cReferencia, D.cCodDest, D.cDireccion, P.cPersCod, " _
    & "ISNULL(P.dPersNacCreac,'') dFecNac From " & sDBPersona & "PersID ID RIGHT JOIN " & sDBPersona & "Persona P INNER JOIN ProductoPersona PG " _
    & "INNER JOIN Producto PD INNER JOIN Giro G INNER JOIN " _
    & "(Select PD.cCtaCod, cDestinatario = Convert(Varchar(120),P.cPersNombre), cDireccion = P.cPersDireccDomicilio, " _
    & "cReferencia = Convert(Varchar(255),ISNULL(ID.cPersIDnro,'')), cCodDest = P.cPersCod From " _
    & sDBPersona & "PersID ID RIGHT JOIN " & sDBPersona & "Persona P INNER JOIN ProductoPersona PG INNER JOIN Producto PD ON " _
    & "PG.cCtaCod = PD.cCtaCod ON P.cPersCod = PG.cPersCod ON ID.cPersCod = P.cPersCod Where PD.nPrdEstado NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ") " _
    & "And PG.nPrdPersRelac = " & gGiroRelPersDestinatario & " " _
    & "UNION " _
    & "Select PD.cCtaCod, D.cDestinatario, cDireccion = '', D.cReferencia, cCodDest = '' From GiroDestinatario D INNER JOIN Producto PD ON " _
    & "D.cCtaCod = PD.cCtaCod Where PD.nPrdEstado NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ")) D ON G.cCtaCod = D.cCtaCod ON " _
    & "PD.cCtaCod = G.cCtaCod ON PG.cCtaCod = PD.cCtaCod ON P.cPersCod = PG.cPersCod ON ID.cPersCod = P.cPersCod INNER JOIN " & sDBComunes & "Agencias A ON " _
    & "G.cAgenciaDest = A.cAgeCod INNER JOIN " & sDBComunes & "Constante K ON G.nPrdCtaTpo = K.nConsValor " _
    & "Where PD.nPrdEstado NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ") And PG.nPrdPersRelac = " & gGiroRelPersRemitente & " And " _
    & "K.nConsCod = " & gProductoCuentaTipo & " And G.cCtaCod = '" & sCuenta & "' Order by cDocID"
rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetGiroDatos = rs
Set rs = Nothing
End Function

Public Function GetGiroPendiente(ByVal sCodage As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient

sSQL = "Select PD.cCtaCod, PD.dPrdEstado, LTRIM(RTRIM(K.cConsDescripcion)) cTipo, PD.nSaldo, " _
    & "P.cPersNombre cRemitente, D.cDestinatario, LTRIM(RTRIM(A.cAgeDescripcion)) cAgencia, " _
    & "D.cReferencia, D.cCodDest From " & sDBPersona & "Persona P " _
    & "INNER JOIN ProductoPersona PG INNER JOIN Producto PD INNER JOIN Giro G INNER JOIN " _
    & "(Select PD.cCtaCod, cDestinatario = Convert(Varchar(120),P.cPersNombre), " _
    & "cReferencia = Convert(Varchar(255),ISNULL(ID.cPersIDnro,'')), cCodDest = P.cPersCod " _
    & "From " & sDBPersona & "PersID ID RIGHT JOIN " & sDBPersona & "Persona P INNER JOIN ProductoPersona PG INNER JOIN Producto PD " _
    & "ON PD.cCtaCod = PG.cCtaCod ON P.cPersCod = PG.cPersCod ON ID.cPersCod = P.cPersCod " _
    & "Where PD.nPrdEstado NOT IN (1300,1400) And PG.nPrdPersRelac = " & gGiroRelPersDestinatario & " UNION " _
    & "Select PD.cCtaCod, D.cDestinatario, D.cReferencia, cCodDest = '' From GiroDestinatario D " _
    & "INNER JOIN Producto PD ON D.cCtaCod = PD.cCtaCod Where PD.nPrdEstado NOT IN (1300,1400)) D " _
    & "ON G.cCtaCod = D.cCtaCod ON PD.cCtaCod = G.cCtaCod ON PD.cCtaCod = PG.cCtaCod ON P.cPersCod = PG.cPersCod " _
    & "INNER JOIN " & sDBComunes & "Agencias A ON G.cAgenciaDest = A.cAgeCod INNER JOIN " & sDBComunes & "Constante K ON G.nPrdCtaTpo = K.nConsValor " _
    & "Where PD.nPrdEstado NOT IN (1300,1400) And PG.nPrdPersRelac = " & gGiroRelPersRemitente & " And K.nConsCod = " & gProductoCuentaTipo & "  and   G.cagenciadest='" & sCodage & "'" _
    & " Order by PD.dPrdEstado"

rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetGiroPendiente = rs
Set rs = Nothing
End Function

Public Function GetPersonasExoLavDinero() As ADODB.Recordset
Dim rsLav As ADODB.Recordset
sSQL = "Select L.cPersCod, RTRIM(P.cPersNombre) Nombre, (K.cConsDescripcion + SPACE(50) + Convert(Varchar(2),L.nEstado)) Estado, " _
    & "RTRIM(L.cComentario) Comentario From Persona P INNER JOIN PersExoLavDinero L ON P.cPersCod = L.cPersCod INNER JOIN Constante K ON " _
    & "L.nEstado = K.nConsValor Where L.cMovNro IN (Select MAX(L1.cMovNro) From PersExoLavDinero L1 Where L1.cPersCod = L.cPersCod) " _
    & "And K.nConsCOd = " & gPersEstLavDinero

Set rsLav = New ADODB.Recordset
rsLav.CursorLocation = adUseClient
rsLav.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsLav.ActiveConnection = Nothing
Set GetPersonasExoLavDinero = rsLav
Set rsLav = Nothing
End Function



Public Function GetPersonaExoneradaLavDinero(ByVal sPersona As String) As ADODB.Recordset
Dim rsLav As ADODB.Recordset
sSQL = "Select L.cPersCod, RTRIM(P.cPersNombre) Nombre, (K.cConsDescripcion + SPACE(50) + Convert(Varchar(2),L.nEstado)) Estado, L.nEstado, " _
    & "RTRIM(L.cComentario) Comentario From Persona P INNER JOIN PersExoLavDinero L ON P.cPersCod = L.cPersCod INNER JOIN Constante K ON " _
    & "L.nEstado = K.nConsValor Where L.cMovNro IN (Select MAX(L1.cMovNro) From PersExoLavDinero L1 Where L1.cPersCod = L.cPersCod) " _
    & "And K.nConsCOd = " & gPersEstLavDinero & " And L.cPersCod = '" & sPersona & "'"

Set rsLav = New ADODB.Recordset
rsLav.CursorLocation = adUseClient
rsLav.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsLav.ActiveConnection = Nothing
Set GetPersonaExoneradaLavDinero = rsLav
Set rsLav = Nothing
End Function

Public Function GetInfoRP(ByVal sAgencia As String) As ADODB.Recordset
Dim rsrp As ADODB.Recordset

sSQL = " select R.CUSER , P.CPERSNOMBRE , R.CPERSCOD "
sSQL = sSQL & " from rrhh R "
sSQL = sSQL & " JOIN PERSONA P ON  P.CPERSCOD=R.CPERSCOD "
sSQL = sSQL & " where nrhestado=201 and crhcod like 'E%' "
sSQL = sSQL & " and careacod in ('026','066') and cagenciaactual='" & sAgencia & "'"
sSQL = sSQL & " ORDER BY P.CPERSNOMBRE "

Set rsrp = New ADODB.Recordset
rsrp.CursorLocation = adUseClient
rsrp.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsrp.ActiveConnection = Nothing
Set GetInfoRP = rsrp
Set rsrp = Nothing
End Function


Public Function GetInfoActa(ByVal sCodUsu As String) As ADODB.Recordset
Dim rsrp As ADODB.Recordset

sSQL = " select FECHA=CONVERT(CHAR(10),FechaIncautacion,103),PORTADOR=p.CPERSNOMBRE , NroActa ,FechaReg=CAST(LEFT(CMOVNRO,8) AS DATETIME), CodPortador "
sSQL = sSQL & " from actasincautacion a "
sSQL = sSQL & " join persona p on p.cperscod= a.CodPortador "
sSQL = sSQL & " where right(cmovnro,4)='" & sCodUsu & "'"
sSQL = sSQL & " ORDER BY NROACTA DESC "

Set rsrp = New ADODB.Recordset
rsrp.CursorLocation = adUseClient
rsrp.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsrp.ActiveConnection = Nothing
Set GetInfoActa = rsrp
Set rsrp = Nothing
End Function

Public Function GetInfoActaCompleta(ByVal sCodUsu As String, ByVal sNroActa As String) As ADODB.Recordset
Dim rsrp As ADODB.Recordset

sSQL = " Select FECHA=FechaIncautacion,PORTADOR=p.CPERSNOMBRE,DNI=(SELECT TOP 1 CPERSIDNRO "
sSQL = sSQL & " FROM PERSID I WHERE I.CPERSCOD=A.CodPortador AND I.CPERSIDTPO=1) , VERIFICADOR=r.cuser+': '+P1.cpersnombre , NroActa,CodPortador,Codusuverificador,a.cmoneda,FechaReg=cast(left(a.cmovnro,8) as datetime) "
sSQL = sSQL & " from actasincautacion a "
sSQL = sSQL & " join persona p on p.cperscod= a.CodPortador "
sSQL = sSQL & " join persona p1 on p1.cperscod= a.codusuverificador "
sSQL = sSQL & " join rrhh r on r.cperscod= a.codusuverificador "
sSQL = sSQL & " where codusudetector='" & sCodUsu & "' and nroacta='" & sNroActa & "'"
sSQL = sSQL & " ORDER BY cast(NROACTA as bigint) DESC "

Set rsrp = New ADODB.Recordset
rsrp.CursorLocation = adUseClient
rsrp.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsrp.ActiveConnection = Nothing
Set GetInfoActaCompleta = rsrp
Set rsrp = Nothing
End Function

Public Function GetInfoDetActa(ByVal sNroActa As String) As ADODB.Recordset
Dim rsrp As ADODB.Recordset

sSQL = " select serienum,fechaemision=CONVERT(CHAR(10),FECHAEMISION,103),denominacion "
sSQL = sSQL & " From detactaincautacion "
sSQL = sSQL & " Where nroacta = '" & sNroActa & "' order by nitem"

Set rsrp = New ADODB.Recordset
rsrp.CursorLocation = adUseClient
rsrp.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsrp.ActiveConnection = Nothing
Set GetInfoDetActa = rsrp
Set rsrp = Nothing
End Function

Public Function GrabaActa(ByVal CodPortador As String, ByVal FechaIncautacion As Date, ByVal CodUsuDetector As String, ByVal COdUsuVerificador As String, ByVal cMovNro As String, ByVal cMoneda As String, ByVal RSTMP As ADODB.Recordset, ByRef NroActa As String) As Boolean
Dim rsrp As ADODB.Recordset, sSQL As String
Dim sNroActa As String, i As Integer

i = 0
On Error GoTo ERRMEN
dbCmact.BeginTrans

sSQL = " select NUEVONUMERO=isnull(MAX(NROACTA),0)+1 "
sSQL = sSQL & " From actaSincautacion "
'ssql = ssql & " Where nroacta = '" & sNroActa & "'"
Set rsrp = New ADODB.Recordset
rsrp.CursorLocation = adUseClient
rsrp.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
If Not rsrp.EOF Then

    'sNroActa = String(10 - Len(rsrp!NuevoNumero), "0") & rsrp!NuevoNumero
    
    sNroActa = rsrp!NuevoNumero
    sSQL = " Insert Into  ActasIncautacion "
    sSQL = sSQL & " values('" & sNroActa & "','" & CodPortador & "','" & Format(FechaIncautacion, "yyyy-mm-dd") & "','" & CodUsuDetector & "','" & COdUsuVerificador & "','" & cMovNro & "','" & cMoneda & "') "
    dbCmact.Execute sSQL
        
    While Not RSTMP.EOF
        i = i + 1
        sSQL = " Insert Into  DetActaIncautacion "
        sSQL = sSQL & " values('" & sNroActa & "'," & i & ",'" & RSTMP.Fields("SERIE Y NUMERACION").Value & "', '" & Format(Left(RSTMP.Fields(1).Value, 10), "mm/dd/yyyy") & "'," & CCur(RSTMP.Fields(2).Value) & ") "
        dbCmact.Execute sSQL
            
            RSTMP.MoveNext
    Wend
    NroActa = sNroActa
    GrabaActa = True
    
Else

    GrabaActa = False
    
End If

dbCmact.CommitTrans
Set rsrp.ActiveConnection = Nothing
'Set GrabaActa = rsrp
Set rsrp = Nothing

Exit Function
ERRMEN:
'Err.Raise Err.Number, "", Err.Description
MsgBox "No se guardó la información. Error:" & Err.Number & vbCrLf & "Porfavor, Inténtelo otra vez", vbOKOnly + vbInformation, "AVISO"
dbCmact.RollbackTrans

End Function

Public Function GetPersonaHistExoLavDinero(ByVal sPersCod As String) As ADODB.Recordset
Dim rsLav As ADODB.Recordset
sSQL = "Select (SUBSTRING(L.cMovNro,7,2) + '/' + SUBSTRING(L.cMovNro,5,2) + '/' + SUBSTRING(L.cMovNro,1,4)) Fecha, " _
    & "(K.cConsDescripcion + SPACE(50) + Convert(Varchar(2),L.nEstado)) Estado, " _
    & "RTRIM(L.cComentario) Comentario, L.cPersCod, L.cMovNro From PersExoLavDinero L INNER JOIN Constante K ON " _
    & "L.nEstado = K.nConsValor Where K.nConsCOd = " & gPersEstLavDinero & " And L.cPersCod = '" & sPersCod & "' " _
    & "Order by L.cMovNro"

Set rsLav = New ADODB.Recordset
rsLav.CursorLocation = adUseClient
rsLav.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsLav.ActiveConnection = Nothing
Set GetPersonaHistExoLavDinero = rsLav
Set rsLav = Nothing
End Function

Public Function GetTitularExoLavDinero(ByVal sCuenta As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient

sSQL = "Select PP.cPersCod, L.cPersCod PersonaExo From ProductoPersona PP LEFT JOIN " _
    & "(Select L.cPersCod, L.nEstado From PersExoLavDinero L Where cMovNro IN (Select MAX(cMovNro) From " _
    & "PersExoLavDinero L1 Where L1.cPersCod = L.cPersCod And L1.nEstado = " & gPersLavDinSinControl & ") " _
    & "And L.nEstado = " & gPersLavDinSinControl & ") L ON PP.cPersCod = L.cPersCod Where " _
    & "PP.nPrdPersRelac = " & gCapRelPersTitular & " And PP.cCtaCod = '" & sCuenta & "'"

rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetTitularExoLavDinero = rs
Set rs = Nothing
End Function

Public Sub AgregaPersExoLavDinero(ByVal sPersCod As String, ByVal sMovNro As String, _
    ByVal nEstado As PersEstLavDinero, ByVal sComentario As String)
sSQL = "Insert PersExoLavDinero (cPersCod,cMovNro,nEstado,cComentario) " _
    & "Values ('" & sPersCod & "','" & sMovNro & "'," & nEstado & ",'" & sComentario & "')"
dbCmact.Execute sSQL
End Sub

Public Sub ActualizaPersExoLavDinero(ByVal sPersCod As String, ByVal sMovNro As String, _
    ByVal sComentario As String, ByVal sMovNroAnt As String)
sSQL = "Update PersExoLavDinero Set cComentario = '" & sComentario & "', cMovNro = '" & sMovNro & "' " _
    & "Where cPersCod = '" & sPersCod & "' And cMovNro = '" & sMovNroAnt & "'"
dbCmact.Execute sSQL
End Sub

Public Sub ActualizaChequeFechaValorizacion(ByVal sNroDoc As String, ByVal sPersCod As String, _
        ByVal sIFTpo As String, ByVal dNuevaFecha As Date)
sSQL = "Update DocRec Set dValorizaRef = '" & Format$(dNuevaFecha, "mm/dd/yyyy") & "' " _
    & "Where nTpoDoc = " & TpoDocCheque & " And cNroDoc = '" & sNroDoc & "' And cPersCod = '" & sPersCod & "' " _
    & "And cIFTpo = '" & sIFTpo & "'"
dbCmact.Execute sSQL
End Sub

Public Sub ActualizaSaldoAnteriorAho(ByVal sCuenta As String, ByVal nSaldoAnt As Double)
sSQL = "Update CaptacAhorros Set nSaldoAnterior = " & nSaldoAnt & "  WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSQL
End Sub

Public Sub ActualizaAbonoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)

If bActExtracto Then
    sSQL = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSQL = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If
dbCmact.Execute sSQL
sSQL = "Update Captaciones Set nSaldoDisp = nSaldoDisp + " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
    & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSQL
End Sub

Public Sub ActualizaCargoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)

If bActExtracto Then
    sSQL = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSQL = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If
dbCmact.Execute sSQL
sSQL = "Update Captaciones Set nSaldoDisp = nSaldoDisp - " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
    & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSQL
End Sub

Public Sub AgregaDocRecEstado(ByVal sNroDoc As String, sPersCod As String, _
        ByVal sTipoIF As String, ByVal nEstado As ChequeEstado, ByVal sMovNro As String, Optional psCtaCheque As String = "")
        
sSQL = "INSERT INTO DocRecEst(nTpoDoc,cNroDoc,cPersCod,cIFTpo,nEstado,cMovNro,CIFcta) " _
     & "VALUES (" & TpoDocCheque & ", '" & sNroDoc & "','" & sPersCod & "','" & Format(sTipoIF, "00") & "'," _
     & nEstado & ",'" & sMovNro & "','" & psCtaCheque & "')"
dbCmact.Execute sSQL
End Sub

Public Sub ActualizaEstadoCheque(ByVal sNroDoc As String, sPersCod As String, _
        ByVal sTipoIF As String, ByVal nEstado As ChequeEstado, Optional dFechaVal As Date)
        
If dFechaVal <> CDate("12:00:00 AM") Then
    sSQL = "Update DocRec Set nEstado = " & nEstado & ", dValorizacion = '" & Format$(dFechaVal, "mm/dd/yyyy") & "' " _
        & "Where nTpoDoc = " & TpoDocCheque & " And cNroDoc = '" & sNroDoc & "' And cPersCod = '" & sPersCod & "' " _
        & "And cIFTpo = '" & sTipoIF & "'"
Else
    sSQL = "Update DocRec Set nEstado = " & nEstado & " " _
        & "Where nTpoDoc = " & TpoDocCheque & " And cNroDoc = '" & sNroDoc & "' And cPersCod = '" & sPersCod & "' " _
        & "And cIFTpo = '" & sTipoIF & "'"
End If
dbCmact.Execute sSQL
End Sub

Public Sub ActualizaFechaRenovacionPF(ByVal sCuenta As String, ByVal dAperturaChq As Date)
sSQL = "Update CaptacPlazoFijo Set dRenovacion = '" & Format$(dAperturaChq, "mm/dd/yyyy") & "' " _
    & "Where cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSQL
End Sub

Public Sub ActualizaSaldoDisponiblePF(ByVal sCuenta As String, ByVal dAperturaChq As Date, ByVal nMonto As Double)
sSQL = "Update Captaciones Set nSaldoDisp = " & nMonto & ", dUltCierre = '" & Format$(DateAdd("d", -1, dAperturaChq), "mm/dd/yyyy") & "' " _
    & "Where cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSQL
End Sub
        
Public Sub ActualizaSaldoRetiroCTS(ByVal sCuenta As String, ByVal nSaldoRet As Double, nIntSaldo As Double)
sSQL = "Update CaptacCTS Set nSaldRetiro = " & nSaldoRet & ", nIntSaldo = " & nIntSaldo & " " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSQL
End Sub

Public Sub AgregaMovCap(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nMonto As Double, _
        ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double)
sSQL = "INSERT MovCap (nMovNro,cOpeCod,cCtaCod,nMonto,nSaldoDisponible,nSaldoContable) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nMonto & "," & nSaldoDisp & "," & nSaldoCnt & ")"
dbCmact.Execute sSQL
End Sub

Public Sub AgregaMovCapDet(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nConcepto As CaptacConcepto, ByVal nMonto As Double)
        
sSQL = "INSERT MovCapDet (nMovNro,cOpeCod,cCtaCod,nConceptoCod,nMonto) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nConcepto & "," & nMonto & ")"
dbCmact.Execute sSQL
End Sub

Public Function GetPorcentajeRetiroCTS(ByVal sCuenta As String, ByVal sNroDoc As String, ByVal sPersCod As String, _
            ByVal sIFTpo As String) As Double
Dim rsChq As ADODB.Recordset
sSQL = "Select P.nPorcentajeAbono From MovParamCTS P INNER JOIN Mov M INNER JOIN DocRecCapta C ON " _
    & "M.nMovNro = C.nMovNro ON P.nMovNro = M.nMovNro Where C.cNroDoc = '" & sNroDoc & "' And " _
    & "C.cPersCod = '" & sPersCod & "' And C.cIFTpo = '" & sIFTpo & "' And C.cCtaCod = '" & sCuenta & "'"

Set rsChq = New ADODB.Recordset
rsChq.CursorLocation = adUseClient
rsChq.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsChq.ActiveConnection = Nothing
GetPorcentajeRetiroCTS = rsChq("nPorcentajeAbono")
rsChq.Close
Set rsChq = Nothing
End Function

Public Function GetProvisionCTS(ByVal nMovNro As Long)
Dim rsChq As ADODB.Recordset
sSQL = "Select MCD.nConceptoCod From Mov M JOIN MovCap MC JOIN MovCapDet MCD ON MC.nMovNro = MCD.nMovNro " _
    & "And MC.cCtaCod = MCD.cCtaCod And MC.cOpeCod = MCD.cOpeCod ON M.nMovNro = MC.nMovNro And " _
    & "M.cOpeCod = MC.cOpeCod Where M.nMovNro = " & nMovNro & " And MCD.nConcepto = " & gConcProvision & " "

Set rsChq = New ADODB.Recordset
rsChq.CursorLocation = adUseClient
rsChq.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsChq.ActiveConnection = Nothing
GetProvisionCTS = rsChq("nConceptoCod")
rsChq.Close
Set rsChq = Nothing
End Function

Public Function GetChequeCuenta(ByVal sNroDoc As String, ByVal sPersCod As String, ByVal sIFTpo As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSQL = "Select cCtaCod, nMonto From DocRecCapta Where cNroDoc = '" & sNroDoc & "' And " _
    & "cPersCod = '" & sPersCod & "' And cIFTpo = '" & sIFTpo & "' Order by cCtaCod"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetChequeCuenta = rs
Set rs = Nothing
End Function

Public Function AgregaGiro(ByVal sCuenta As String, ByVal dApertura As Date, ByVal nMonto As Double, _
    ByVal nTipoGiro As ProductoCuentaTipo, ByVal sAgenciaDest As String)
sSQL = "Insert Producto (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc) " _
    & "Values ('" & sCuenta & "',0," & nMonto & "," & gCapEstActiva & ",'" & Format$(dApertura & " " & Time, "mm/dd/yyyy hh:mm:ss") & "',0)"
dbCmact.Execute sSQL
sSQL = "Insert Giro (cCtaCod,dVigencia,nPrdCtaTpo,cAgenciaDest) " _
    & "Values ('" & sCuenta & "','" & Format$(dApertura, "mm/dd/yyyy") & "'," & nTipoGiro & ",'" & sAgenciaDest & "')"
dbCmact.Execute sSQL
End Function

Public Sub AgregaProdPersGiro(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As GiroRelacPersona)
sSQL = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " _
    & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ")"
dbCmact.Execute sSQL
End Sub

Public Sub AgregaGiroDestinatario(ByVal sCuenta As String, ByVal sDestinatario As String, _
        ByVal sReferencia As String)
sSQL = "Insert GiroDestinatario (cCtaCod,cDestinatario,cReferencia) " _
    & "Values ('" & sCuenta & "','" & sDestinatario & "','" & sReferencia & "')"
dbCmact.Execute sSQL
End Sub

Public Sub ActualizaEstadoProducto(ByVal sCuenta As String, ByVal nEstado As CaptacEstado, _
        ByVal dFecha As Date)
'Actualiza el ultimo estado a la tabla de producto
sSQL = "Update Producto Set nPrdEstado = " & nEstado & ", " _
    & "dPrdEstado = '" & Format$(dFecha & " " & Time, "mm/dd/yyyy hh:mm:ss") & "' WHERE " _
    & "cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSQL
End Sub

' NUEVO - CMCPL
Public Sub AgregaServSATTributo(ByVal sCodTributo As String, ByVal sTributo As String, ByVal sAno As String, _
        ByVal sRecibo As String, ByVal sDigCheq As String, ByVal sPeriodo As String, ByVal nValDeuda As Double, _
        ByVal nValDerEmis As Double, ByVal nValAjuste As Double, ByVal nValIntMor As Double, ByVal dFecVenc As Date, _
        ByVal sNombre As String, ByVal sCuenta As String, ByVal sContrib As String, ByVal nValGastos As Double, _
        ByVal nValCostas As Double, ByVal sCodSis As String, ByVal nTotal As Double)

sSQL = "INSERT CapServSATTributos (cCodTributo,cTributo,cAno,cRecibo,cDigCheq,cPeriodo,nValDeuda,nValDerEmis,nValAjuste, " _
    & "nValIntMor , dFecVenc, cNombre, cCuenta, cContrib, nValGastos, nValCostas, cCodSis, nTotal)  " _
    & "VALUES ('" & sCodTributo & "','" & sTributo & "','" & sAno & "','" & sRecibo & "','" & sDigCheq & "', " _
    & "'" & sPeriodo & "'," & nValDeuda & "," & nValDerEmis & "," & nValAjuste & ", " & nValIntMor & ", " _
    & "'" & Format$(dFecVenc, "mm/dd/yyyy") & "','" & sNombre & "','" & sCuenta & "','" & sContrib & "', " _
    & nValGastos & "," & nValCostas & ",'" & sCodSis & "'," & nTotal & ")"
dbCmact.Execute sSQL
End Sub
' NUEVO - CMCPL
Public Sub AgregaServSATpapeletas(ByVal sCodTributo As String, ByVal sTributo As String, ByVal sAno As String, _
        ByVal sPlaca As String, ByVal sDigCheq As String, ByVal sPeriodo As String, ByVal nImporte As Double, _
        ByVal nValReincide As Double, ByVal nValGastAdm As Double, ByVal nValCostas As Double, ByVal dFecVenc As Date, _
        ByVal dFecInf As Date, ByVal dFecLect As String, ByVal dFecNoti As String, ByVal sPapeleta As String, _
        ByVal sTipo As String, ByVal sFalta As String, ByVal sCodAux As String, ByVal sLicCond As String)

sSQL = "INSERT CapServSATPapeletas (cCodTributo,cTributo,cAno,cPlaca,cDigCheq,cPeriodo,nImporte,nValReincide,nValGastAdm, " _
    & "nValCostas , dFecVenc, dFecInf, dFecLect, dFecNoti, sPapeleta, sTipo, sFalta, sCodAux, sLicCond )  " _
    & "VALUES ('" & sCodTributo & "','" & sTributo & "','" & sAno & "','" & sPlaca & "','" & sDigCheq & "', " _
    & "'" & sPeriodo & "'," & nImporte & "," & nValReincide & "," & nValGastAdm & ", " & nValCostas & ", " _
    & "'" & Format$(dFecVenc, "mm/dd/yyyy") & "', '" & Format$(dFecInf, "mm/dd/yyyy") & "', " _
    & "'" & Format$(dFecLect, "mm/dd/yyyy") & "', '" & Format$(dFecNoti, "mm/dd/yyyy") & "', " _
    & "'" & sPapeleta & "','" & sTipo & "','" & sFalta & "','" & sCodAux & "','" & sLicCond & "')"
    
dbCmact.Execute sSQL
End Sub
' NUEVO - CMCPL
Public Sub EliminaServSATTributo()
sSQL = "TRUNCATE TABLE CapServSATTributos"
dbCmact.Execute sSQL
End Sub
' NUEVO - CMCPL
Public Sub EliminaServSATPapeletas()
'sSql = "TRUNCATE TABLE CapServSATPapeletas"
'dbCmact.Execute sSql
End Sub
' NUEVO - CMCPL
Public Function GetServSATPapeletaLincencia(ByVal sLicencia As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSQL = "Select cPlaca, nImporte, nValGastAdm, nValCostas, dFecInf, sPapeleta, sTipo, sFalta, " _
    & "sLicCond From CapServSATPapeletas Where sLicCond = '" & sLicencia & "' Order by dFecInf"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetServSATPapeletaLincencia = rs
Set rs = Nothing
End Function
' NUEVO - CMCPL
Public Function GetServSATPapeletaPlaca(ByVal sPlaca As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSQL = "Select cPlaca, nImporte, nValGastAdm, nValCostas, dFecInf, sPapeleta, sTipo, sFalta, " _
    & "sLicCond From CapServSATPapeletas Where cPlaca = '" & sPlaca & "' Order by dFecInf"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetServSATPapeletaPlaca = rs
Set rs = Nothing
End Function
' NUEVO - CMCPL
Public Function GetServSATPapeletaNPapeleta(ByVal sPapeleta As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSQL = "Select cPlaca, nImporte, nValGastAdm, nValCostas, dFecInf, sPapeleta, sTipo, sFalta, " _
    & "sLicCond From CapServSATPapeletas Where sPapeleta = '" & sPapeleta & "' Order by dFecInf"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetServSATPapeletaNPapeleta = rs
Set rs = Nothing
End Function
' NUEVO - CMCPL
Public Function GetServSATRecibo(ByVal sCuenta As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSQL = "Select cCodtributo, cnombre, cTributo, cAno, cRecibo, nValDeuda, nValDerEmis, nValCostas, nValAjuste, " _
    & "nValIntMor,cCuenta,cContrib, nValGastos From CapServSATtributos Where cCuenta = '" & sCuenta & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetServSATRecibo = rs
Set rs = Nothing
End Function
' NUEVO - CMCPL
Public Function GetServSATDerecho(ByVal sCodigo As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSQL = "Select cdistrito, ccodDerecho, cTipValor, cDescrip, nValDerecho From CapServicioDerecho Where ccodDerecho= '" & sCodigo & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetServSATDerecho = rs
Set rs = Nothing
End Function
' NUEVO - CMCPL CRSF 08/10
Public Function GetSatMontoFecha(ByVal dInicio As String, ByVal dFin As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSQL = "Select a.nMovnro,substring(a.cMovNro,22,4) USUA, a.cOpeCod, b.nmonto, " _
    & " a.cMovdesc,substring(cmovNro,7,2)+'/'+substring(cMovNro,5,2)+'/'+substring(cMovNro,1,4) fecha" _
    & " From mov  a, movservicios b Where a.nmovnro = b.nmovnro  and  cOpeCod in ('300106','300105')  and " _
    & "substring(cmovNro,7,2)+'/'+substring(cMovNro,5,2)+'/'+substring(cMovNro,1,4) = '" & dInicio & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetSatMontoFecha = rs
Set rs = Nothing
End Function

' NUEVO - CMCPL CRSF 10/10
Public Function GetSatMontoFechaDis(ByVal dInicio As String, ByVal dFin As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSQL = "Select a.cOpeCod,b.nPrdConcepto,sum(b.nMonto) nMonto From mov a, movserviciosdet b " _
        & " where a.nmovnro = b.nmovnro  and  a.cOpeCod in ('300106','300105')  and " _
    & "substring(cmovNro,7,2)+'/'+substring(cMovNro,5,2)+'/'+substring(cMovNro,1,4) = '" & dInicio & "'" _
    & " group by  a.cOpeCod,b.nPrdConcepto order by a.cOpeCod,b.nPrdConcepto"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetSatMontoFechaDis = rs
Set rs = Nothing
End Function
'NUEVO - CMCPL CRSF 11/06 - Monto de los Parametros de Servicios para Distrib
Public Function GetSatDistribucion(ByVal sParam As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSQL = "Select cCtacodAbono,nComision from CapServicioParam " _
        & "Where nservcod = 300000 And nTipoComision = '" & sParam & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetSatDistribucion = rs
Set rs = Nothing
End Function

'NUEVO - CMCPL CRSF 11/06 - Suma de Los Totales
Public Function GetSatSuma(ByVal sParam As String) As Double
Dim rs As ADODB.Recordset
sSQL = "Select sum(b.nMonto) nMonto from mov a, movserviciosdet b " _
        & "where a.nMovNro = b.nMovNro and cOpeCod in ('300105','300106') and nPrdConcepto ='" & sParam & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rs.ActiveConnection = Nothing
GetSatSuma = rs("nMonto")
rs.Close
Set rs = Nothing
End Function

Public Function GetNumSolcitudCapTasaEspecial() As Long
Dim rs As ADODB.Recordset
sSQL = "Select nNumSolicitud = ISNULL(MAX(nNumSolicitud),0) + 1 From CapTasaEspecial"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
GetNumSolcitudCapTasaEspecial = rs("nNumSolicitud")
rs.Close
Set rs = Nothing
End Function

Public Sub AgregaCapTasaEspecial(ByVal nNumSolicitud As Integer, ByVal sPersona As String, ByVal nProd As Producto, ByVal nMon As Moneda, _
        ByVal nEstado As Integer, ByVal sMovNro As String, ByVal nTasa As Double, ByVal sComentario As String, _
        ByVal nMonto As Double, Optional ByVal sCuenta As String = "", Optional ByVal nPlazo As Integer = 0, Optional ByVal nExtorno As Integer = 99, Optional ByVal bPermanente As Boolean)
        
Dim rs As ADODB.Recordset

If nExtorno <> 99 Then
    sSQL = " Insert CapTasaEspecial (nNumSolicitud,cPersCod,nProducto,nMoneda,nEstado,cMovNro,nTasa,cComentario,cCtaCod,nMonto,nPlazo,nextorno,bPermanente) " _
         & "  VALUES (" & nNumSolicitud & ",'" & sPersona & "'," & Trim(nProd) & "," & Trim(nMon) & "," & Trim(nEstado) & ",'" & sMovNro & "'," & Trim(nTasa) & ",'" & sComentario & "','" & sCuenta & "'," & Trim(nMonto) & "," & Trim(nPlazo) & "," & nExtorno & "," & IIf(bPermanente, 1, 0) & " ) "

Else
    sSQL = " Insert CapTasaEspecial (nNumSolicitud,cPersCod,nProducto,nMoneda,nEstado,cMovNro,nTasa,cComentario,cCtaCod,nMonto,nPlazo,bpermanente) " _
         & "  VALUES (" & nNumSolicitud & ",'" & sPersona & "'," & Trim(nProd) & "," & Trim(nMon) & "," & Trim(nEstado) & ",'" & sMovNro & "'," & Trim(nTasa) & ",'" & sComentario & "','" & sCuenta & "'," & Trim(nMonto) & "," & Trim(nPlazo) & "," & IIf(bPermanente, 1, 0) & ") "
End If
    
dbCmact.Execute sSQL


End Sub

Public Function GetEstadoTPCta(ByVal sCuenta As String) As ADODB.Recordset
  Dim rs As ADODB.Recordset
  
  
  
  sSQL = " Select  nNumSolicitud,cMovNro,cPersCod,nProducto,nMoneda,cCtaCod,nMonto,nPlazo,ntasa,nestado,cComentario,NEXTORNO,bpermanente=isnull(bpermanente,0) from captasaespecial where nestado=2 and  cctacod='" & sCuenta & "' "
  sSQL = sSQL & " and left(cmovnro,14)+cast(nnumsolicitud as varchar(12)) in (select max(left(c.cmovnro,14)+cast(c.nnumsolicitud as varchar(12))) from captasaespecial c where cctacod='" & sCuenta & "')"
     
  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseClient
  rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
  
 
    Set GetEstadoTPCta = rs
 
  Set rs.ActiveConnection = Nothing
  Set rs = Nothing

End Function



Public Function GetUltEstadoTP(ByVal nNumsol As Long, ByVal scperscod As String, ByVal nProd As Integer, ByVal nMon As Integer) As ADODB.Recordset
  Dim rs As ADODB.Recordset
  
  sSQL = "Select  nNumSolicitud,cMovNro,cPersCod,nProducto,nMoneda,cCtaCod,nMonto,nPlazo,ntasa,nestado,cComentario,NEXTORNO,bpermanente=isnull(bpermanente,0) from captasaespecial where nnumsolicitud=" & nNumsol & " and nproducto=" & nProd & " and cperscod='" & scperscod & "' and nmoneda=" & nMon
  sSQL = sSQL & " and left(cmovnro,14)+cast(nnumsolicitud as varchar(12)) in (select max(left(c.cmovnro,14)+cast(c.nnumsolicitud as varchar(12))) from captasaespecial c where  nnumsolicitud=" & nNumsol & " and nproducto=" & nProd & " and nmoneda=" & nMon & ")"
  
   
  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseClient
  rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
  Set rs.ActiveConnection = Nothing
  Set GetUltEstadoTP = rs
  Set rs = Nothing
 

End Function




Public Function GetDatosTasaEspecialSol() As ADODB.Recordset
Dim rs As ADODB.Recordset


sSQL = " SELECT P.cPersNombre, cProducto = K.cConsDescripcion, nTasa=case when c.nestado<>3  then  dbo.ConvierteTNAaTEA(c.nTasa) else c.ntasa  end  , cMoneda = CASE WHEN c.nMoneda = 1 THEN 'S/.' ELSE 'US$' END, "
sSQL = sSQL & " c.nMonto , c.cComentario, c.cPersCod, c.nMoneda, c.nPlazo, c.cMovnro, c.nProducto, c.nNumSolicitud "
sSQL = sSQL & " FROM  CAPTASAESPECIAL c "
sSQL = sSQL & " JOIN Persona P ON p.cPersCod = c.cPersCod "
sSQL = sSQL & " Left Join "
sSQL = sSQL & " Constante K ON k.nConsValor = c.nProducto and k.nconscod=" & gProducto
sSQL = sSQL & " Where "
sSQL = sSQL & " left(c.cmovnro,14)+cast(c.nnumsolicitud as varchar(12))  in "
sSQL = sSQL & " ( select max(left(c.cmovnro,14))+cast(c.nnumsolicitud as varchar(12)) from CAPTASAESPECIAL c "
sSQL = sSQL & "   group by c.nnumsolicitud  ) "
sSQL = sSQL & "   and (nestado=0 or (nestado=3 and nextorno=1) or (nestado=3 and nextorno=4)) "
sSQL = sSQL & " order by c.nNumSolicitud "


'sSql = " select * from captasaespecial c "
'sSql = sSql & " Left Join "
'sSql = sSql & " captasaespecial c1 on c1.nNumsolicitud=c.nNumsolicitud and c1.nestado=3 "
'sSql = sSql & " Left Join "
'sSql = sSql & " captasaespecial c2 on c2.nNumsolicitud=c.nNumsolicitud and c2.nestado=4 "
'sSql = sSql & " Left Join "
'sSql = sSql & " captasaespecial c3 on c3.nNumsolicitud=c.nNumsolicitud and c3.nestado=2 "
'sSql = sSql & " Left Join "
'sSql = sSql & " captasaespecial c4 on c4.nNumsolicitud=c.nNumsolicitud and c4.nestado=1 "
'sSql = sSql & " Where c1.nNumSolicitud Is Null And c2.nNumSolicitud Is Null "
'sSql = sSql & " and c3.nnumsolicitud is null AND c4.nnumsolicitud is null "
'sSql = sSql & " and  c.nestado=0 "

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetDatosTasaEspecialSol = rs
Set rs = Nothing


End Function


Public Function GetDatosCapTasaEspecial(ByVal nEstado As Integer) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSQL = "Select P.cPersNombre, cProducto = K.cConsDescripcion, T.nTasa, cMoneda = CASE WHEN T.nMoneda = 1 THEN 'S/.' ELSE 'US$' END, " _
    & "T.nMonto, T.cComentario, T.cPersCod, T.nMoneda, T.nPlazo, T.cMovNro, T.nProducto, T.nNumSolicitud From CapTasaEspecial T " _
    & "JOIN Persona P ON T.cPersCod = P.cPersCod LEFT JOIN Constante K ON T.nProducto = K.nConsValor Where " _
    & "T.cMovNro IN (Select Max(cMovNro) From CapTasaEspecial T1 Where T1.cPersCod = T.cPersCod And " _
    & "T1.nProducto = T1.nProducto And T1.nMoneda = T.nMoneda And T1.nNumSolicitud = T.nNumSolicitud) " _
    & "And T.nEstado = " & nEstado & " And K.nConsCod = " & gProducto
    
    
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetDatosCapTasaEspecial = rs
Set rs = Nothing
End Function

Public Function GetCargosCartasAI(ByVal nTipo As Integer, Filtro1 As String, Filtro2 As String, Filtro3 As String, psCodUser As String) As Boolean
Dim sSQL As String, RSTMP As ADODB.Recordset
  Dim Con As COMConecta.DCOMConecta
  Set Con = New COMConecta.DCOMConecta
  Set RSTMP = New ADODB.Recordset
  Dim ssqladd As String
  

    
  If nTipo = 1 Then

        ssqladd = " WHERE NROCARTA>='" & Filtro1 & "' and NROCARTA<='" & Filtro2 & "'" & IIf(Filtro3 = "", "", IIf(Filtro3 = "11", " and ubigeo like '_11%'", " and ubigeo not like '_11%'"))

  Else
       
        ssqladd = " WHERE LEFT(CMOVNRO,8)>='" & Filtro1 & "' and left(cmovnro,8)<='" & Filtro2 & "'" & IIf(Filtro3 = "", "", IIf(Filtro3 = "11", " and ubigeo like '_11%'", " and ubigeo not like '_11%'"))

  End If
  
  Con.AbreConexion
  
  sSQL = "CREATE TABLE #TEMPCARTA" & psCodUser & "(COD int IDENTITY(1, 1) NOT NULL,DATO VARCHAR(2000),CARTA VARCHAR(10)) "
  Call Con.Ejecutar(sSQL)
  
  sSQL = "  INSERT into #tempcarta" & psCodUser & "(DATO ,CARTA ) "
'  sSQL = sSQL & "  select  dato1= char(13)+ ca.entidad + char(13)"
'  sSQL = sSQL & "  + ltrim(rtrim(ISNULL(U1.CUBIGEODES,''))) + '-' + ltrim(rtrim(ISNULL(U2.CUBIGEODES,''))) + char(13)"
'  sSQL = sSQL & "  + 'CARTA N° ' + ca.nrocarta + '-CMI-S.A-2004/GAH/J' + char(13)"
'  sSQL = sSQL & "  + y.cconsdescripcion + ' ' + ca.cnrodoc + char(13),ca.nrocarta"
'  sSQL = sSQL & "  from (select * from cartasai where cnroref=1 ) ca"
'  sSQL = sSQL & " join (select * from constante where nconscod=2027) k on k.nconsvalor=ca.ntipodelito"
'  sSQL = sSQL & " join (select * from constante where nconscod=2028) y on y.nconsvalor=ca.ntipodoc"
'  sSQL = sSQL & " LEFT JOIN (SELECT * FROM UBIGEO WHERE CUBIGEOCOD LIKE '1%' ) U1 ON U1.CUBIGEOCOD= '1' + SUBSTRING(CA.UBIGEO,2,2)+REPLICATE('0',9)"
'  sSQL = sSQL & " LEFT JOIN (SELECT * FROM UBIGEO WHERE CUBIGEOCOD LIKE '2%' ) U2 ON U2.CUBIGEOCOD= '2' + SUBSTRING(CA.UBIGEO,2,4)+REPLICATE('0',7)"
'  sSQL = sSQL & ssqladd
'  sSQL = sSQL & " ORDER BY CA.NROCARTA"
   sSQL = sSQL & "SELECT DATO,CARTA FROM CARTASMREYNA ORDER BY NRO "

  Call Con.Ejecutar(sSQL)
  
   sSQL = " SELECT J.COD,J.DATO1,J.CARTA1,F.COD,DATO2=ISNULL(F.DATO2,''),ISNULL(F.CARTA2,0) "
   sSQL = sSQL & "     From "
   sSQL = sSQL & "     (SELECT  COD,DATO1=DATO,CARTA1=CARTA FROM #tempcarta" & psCodUser
   sSQL = sSQL & "      WHERE COD%2<>0 ) J "
   sSQL = sSQL & "     Left Join "
   sSQL = sSQL & "     (SELECT  COD,DATO2=DATO,CARTA2=CARTA FROM #tempcarta" & psCodUser
   sSQL = sSQL & "      WHERE COD%2=0  ) F ON F.COD=(J.COD+1) "
'   sSQL = sSQL & "     ORDER BY J.CARTA1,F.CARTA2 "
   RSTMP.Open sSQL, Con.ConexionActiva, adOpenStatic, adLockOptimistic, adCmdText
    
    
    
'*********** Los Reportes se van a generar en otro momento
    
'    If Not RSTMP.EOF Then
'              Dim oRep As DRCargosCARTASAI
'              Set oRep = New DRCargosCARTASAI
'              'Set DRCargosCARTASAI.DataSource = rstmp
'              Set oRep.DataSource = RSTMP
'
'              oRep.Show 1
'              Set oRep = Nothing
'              GetCargosCartasAI = True
'
'    Else
'
'              GetCargosCartasAI = False
'    End If
    
'*******************************************************
    
   sSQL = sSQL & " drop table #tempcarta" & psCodUser
    
  
 Con.CierraConexion
 Set Con = Nothing

End Function
