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

Public dbCmact As ADODB.Connection
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String
Dim sSql As String

Dim oFun As New COMFunciones.FCOMVarPublicas

Public Sub GrabaMovCobranzas(ByVal nMovNro As Long, ByVal cPersCod As String, ByVal cCodigo As String, ByVal nTpoCuenta As Integer, ByVal nMonto As Double, ByVal cReferencia As String)

sSql = "Insert movconvcobranza(nMovNro ,cPersCod ,cCodigo ,nTpoCuenta ,nMonto ,cReferencia)  " _
       & " VALUES (" & nMovNro & ",'" & cPersCod & "','" & cCodigo & "'," & nTpoCuenta & "," & nMonto & ",'" & cReferencia & "' ) "

dbCmact.Execute sSql
    
    
End Sub

Public Function GetnMovNro(ByVal sMovNro As String) As Long
Dim rsMov As ADODB.Recordset
sSql = "Select nMovNro From Mov Where cMovNro = '" & sMovNro & "'"
Set rsMov = New ADODB.Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
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
sSql = "Select nComision, cCtaCodAbono, nTipoComision From " _
    & "CapServicioParam Where nServCod = " & nServicio
rsServ.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsServ.ActiveConnection = Nothing
Set GetServicioParametros = rsServ
Set rsServ = Nothing
End Function

Public Sub ActualizaServicioParametro(ByVal nServicio As CaptacInstServicios, _
        ByVal sCuenta As String, ByVal nComision As Double, ByVal nTipoComision As CapServTipoComision)
sSql = "Update CapServicioParam Set nComision = " & nComision & ", " _
    & "cCtaCodAbono = '" & sCuenta & "', " _
    & "nTipoComision = " & nTipoComision & " " _
    & "Where nServCod = " & nServicio
dbCmact.Execute sSql
End Sub
Public Sub AgregaMovLavDinero(ByVal nMovNro As Long, ByVal sPerscod As String, _
                                Optional ByVal psTitPersLavDinero As String, _
                                Optional ByVal psOrdPersLavDinero As String, _
                                Optional ByVal psReaPersLavDinero As String, _
                                Optional ByVal psBenPersLavDinero As String, _
                                Optional ByVal psVisPersLavDinero As String)
                                
'By Capi 18022008
'Public Sub AgregaMovLavDinero(ByVal nMovNro As Long, ByVal sPersCod As String)
'sSql = "INSERT MovLavDinero (nMovNro,cPersCod) " _
'  & "VALUES (" & nMovNro & ",'" & sPersCod & "')"
                             

sSql = "INSERT MovLavDinero (nMovNro,cPersCod,cPersCodOrd,cPersCodRea,cPersCodBen,cPersCodVisto) " _
    & "VALUES (" & nMovNro & ",'" & psTitPersLavDinero & "','" & psOrdPersLavDinero & "','" & psReaPersLavDinero & "','" & psBenPersLavDinero & "','" & psVisPersLavDinero & "')"
dbCmact.Execute sSql
End Sub

Public Sub AgregaServicioParametro(ByVal nServicio As CaptacInstServicios, _
        ByVal sCuenta As String, ByVal nComision As Double, ByVal nTipoComision As CapServTipoComision)
sSql = "Insert CapServicioParam (nServCod,nComision,cCtaCodAbono,nTipoComision)  " _
    & "VALUES (" & nServicio & "," & nComision & ",'" & sCuenta & "'," & nTipoComision & ")"
dbCmact.Execute sSql
End Sub

Public Function EsReciboCobrado(ByVal nOperacion As CaptacOperacion, ByVal sNumRecibo As String) As Boolean
Dim rsRec As ADODB.Recordset
Set rsRec = New ADODB.Recordset
rsRec.CursorLocation = adUseClient
sSql = "Select M.cMovNro, S.cNumRecibo, S.cCodigo From Mov M INNER JOIN MovServicios S " _
    & "ON M.nMovNro = S.nMovNro Where S.cNumRecibo = '" & sNumRecibo & "' And " _
    & "M.cOpeCod = '" & nOperacion & "'"
rsRec.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
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
sSql = "UPDATE Mov Set nMovEstado = " & nMovEstado & ", nMovFlag = " & nMovFlag & " " _
    & "WHERE nMovNro = '" & nMovNro & "'"
dbCmact.Execute sSql
End Sub

Public Sub AgregaMov(ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
        ByVal sGlosa As String, Optional nMovEstado As MovEstado = gMovEstContabMovContable, _
        Optional nMovFlag As MovFlag = gMovFlagVigente)
    sSql = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
           & "VALUES ('" & sMovNro & "','" & nOperacion & "','" & sGlosa & "'," & nMovEstado & "," & nMovFlag & ")"
    dbCmact.Execute sSql
End Sub

Public Sub AgregaMovRef(ByVal nMovNro As Long, ByVal nMovNroRef As Long)
sSql = "INSERT MovRef (nMovNro,nMovNroRef) " _
    & "VALUES (" & nMovNro & "," & nMovNroRef & ")"
dbCmact.Execute sSql
End Sub
Public Sub AgregaMovServicios(ByVal nMovNro As Long, ByVal sNumRecibo As String, _
       ByVal sCodigo As String, ByVal nMonto As Double, _
        Optional nFlag As CaptacFlagServicios = gCapServFlagRegistrado, _
        Optional nMoneda As Moneda = gMonedaNacional, Optional gcOpeCod As String = "")
    
    sSql = "INSERT MovServicios (nMovNro,cNumRecibo,cCodigo,nMonto,nFlag,nMoneda, cOpeCod) " _
           & "VALUES (" & nMovNro & ",'" & sNumRecibo & "','" & sCodigo & "'," & nMonto & "," & nFlag & "," & nMoneda & ", '" & gcOpeCod & "')"
    dbCmact.Execute sSql
End Sub

Public Sub AgregaMovConvCobranza(ByVal nMovNro As Long, ByVal cPersCod As String, ByVal cCodigo As String, ByVal nTpoCuenta As Integer, ByVal nMonto As Double, ByVal cReferencia As String)
sSql = "INSERT movconvcobranza (nMovNro,cPersCod,cCodigo,nTpoCuenta,nMonto,cReferencia) " _
    & "VALUES (" & nMovNro & ",'" & cPersCod & "','" & cCodigo & "'," & nTpoCuenta & "," & nMonto & ",'" & cReferencia & "')"
dbCmact.Execute sSql
End Sub

Public Sub AgregaMovServiciosDet(ByVal nMovNro As Long, ByVal sNumRecibo As String, _
        ByVal nConcepto As CaptacConcepto, ByVal nMonto As Double, Optional nCuota As Long = 0, Optional gcOpeCod As String = "")

    sSql = "INSERT MovServiciosDet (nMovNro,cNumRecibo,nPrdConcepto,nNroCuota,nMonto, cOpeCod) " _
           & "VALUES (" & nMovNro & ",'" & sNumRecibo & "'," & nConcepto & "," & nCuota & "," & nMonto & ", '" & gcOpeCod & "')"
    dbCmact.Execute sSql
End Sub

Public Function ActualizaFlagServicios(ByVal nMovNro As Long, ByVal nFlag As CaptacFlagServicios)
sSql = "UPDATE MovServicios Set nFlag  = " & nFlag & " WHERE nMovNro = " & nMovNro
dbCmact.Execute sSql
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

sSql = "Select dFecha = CONVERT(Varchar(10),CONVERT(Datetime,LEFT(M.cMovNro,8),103)) + ' ' + " _
    & "SUBSTRING(M.cMovNro,9,2) + ':' + SUBSTRING(M.cMovNro,11,2) + ':' + SUBSTRING(M.cMovNro,13,2), " _
    & "S.cNumRecibo, S.cCodigo, SD.nMonto, A.cAgeDescripcion, RIGHT(M.cMovNro,4) Usuario, M.nMovNro " _
    & "From Mov M INNER JOIN MovServicios S JOIN MovServiciosDet SD ON S.nMovNro = SD.nMovNro And " _
    & "S.cNumRecibo = SD.cNumRecibo ON M.nMovNro = S.nMovNro JOIN Agencias A ON " _
    & "SUBSTRING(M.cMovNro,18,2) = A.cAgeCod Where LEFT(M.cMovNro,8) BETWEEN '" _
    & Format$(dInicio, "yyyymmdd") & "' And '" & Format$(dFin, "yyyymmdd") & "' And M.nMovFlag IN (" & gMovFlagVigente & ") " _
    & "And M.cOpeCod = '" & nServicio & "' And SD.nPrdConcepto = " & gConcCapital & " " & sRef
    
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetMovServicios = rsMov
Set rsMov = Nothing
End Function

Public Sub CreaTablaDBFSedalib(ByVal sNombreArchivo As String)
Dim dbAux As ADODB.Connection
Set dbAux = New ADODB.Connection

sSql = "CREATE TABLE " & sNombreArchivo & " ( "
sSql = sSql & "FACSERNRO N(2,0) NOT NULL, "
sSql = sSql & "FACNRO N(7,0)  NOT NULL, "
sSql = sSql & "HISEMIFEC D(8) NOT NULL, "
sSql = sSql & "HISCOBFEC D(8) NOT NULL, "
sSql = sSql & "HISCORNRO N(5,0) NOT NULL, "
sSql = sSql & "OFICOD N(1,0) NOT NULL, "
sSql = sSql & "OFIAGECOD N(0,0) NOT NULL, "
sSql = sSql & "HISFACANO C(4) NOT NULL, "
sSql = sSql & "HISFACMES C(2) NOT NULL, "
sSql = sSql & "CLIUNICOD C(11) NOT NULL, "
sSql = sSql & "HISCHEDIG N(1,0) NOT NULL, "
sSql = sSql & "HISDIGCHE N(1,0) NOT NULL, "
sSql = sSql & "HISCOBHRS C(8) NOT NULL, "
sSql = sSql & "HISFLAGZ N(0,0) NOT NULL, "
sSql = sSql & "HISACTFEC D(8) NOT NULL, "
sSql = sSql & "HISACTFLA N(0,0) NOT NULL, "
sSql = sSql & "HISCTRFLA N(0,0) NOT NULL, "
sSql = sSql & "HISPAGO N(8,2) NOT NULL, "
sSql = sSql & "HISCOBTIP  C(1) NOT NULL, "
sSql = sSql & "HISEXCCOB N(10,2) NOT NULL,"
sSql = sSql & "HISCODZON N(0,0) NOT NULL, "
sSql = sSql & "HISCODREG N(0,0) NOT NULL, "
sSql = sSql & "HISDIGCOD N(1,0) NOT NULL, "
sSql = sSql & "HISCOBDIA N(11,2) NOT NULL, "
sSql = sSql & "HISGRUCOD N(2,0) NOT NULL, "
sSql = sSql & "HISGRUSUB N(3,0) NOT NULL, "
sSql = sSql & "NFILE  C(2) NOT NULL )"
    
dbAux.Open oFun.gsConnServDBF
dbAux.Execute sSql
dbAux.Close
Set dbAux = Nothing
End Sub

Public Sub AgregaServDBF(ByVal sArchivo As String, ByVal sRecibo As String, _
        ByVal sCodigo As String, ByVal nMonto As Double, ByVal dFecha As Date, _
        ByVal sHora As String)
Dim dbAux As ADODB.Connection
Set dbAux = New ADODB.Connection
sSql = "INSERT INTO " & sArchivo & " (Facsernro,Facnro,hisemifec,Hiscobfec,hiscornro,oficod,ofiagecod,hisfacano,hisfacmes,cliunicod,hischedig,hisdigche,hiscobhrs,hisflagz,hisactfec,hisactfla,hisctrfla,hispago,hiscobtip,hisexccob,hiscodzon,hiscodreg,hisdigcod,hiscobdia,hisgrucod,hisgrusub,nfile) " _
    & "VALUES(" & Left(sRecibo, 3) & "," & Mid(sRecibo, 4, 8) & ",CTOD('" & Format(dFecha, "mm/dd/yyyy") & "'),CTOD('" & Format(dFecha, "mm/dd/yyyy") & "'),0,5,0,'','','" & sCodigo & "',0,0,'" & sHora & "',0," _
    & "CTOD('" & Format(dFecha, "mm/dd/yyyy") & "'),0,0," & Format(nMonto, "##0.00") & ",'',0,0,0,4,0,0,0,'')"
dbAux.Open oFun.gsConnServDBF
dbAux.Execute sSql
dbAux.Close
Set dbAux = Nothing
End Sub
Public Function GetServConvenios2() As ADODB.Recordset
Dim rsConv As ADODB.Recordset
Set rsConv = New ADODB.Recordset
rsConv.CursorLocation = adUseClient
sSql = "Select distinct Codigo=C.cPersCod, Institucion=rtrim(ltrim(P.cPersNombre)) , Tipo=case when c.nconvcod=101 then 'OTROS INGRESOS CON PERSONAS' " _
    & " else 'COLEGIOS' end, c.nconvcod " _
    & "  From CaptacConvenio C INNER JOIN " & sDBPersona & "Persona P " _
    & "ON C.cPersCod = P.cPersCod INNER JOIN " & sDBComunes & "Constante K ON C.nConvCod = K.nConsValor " _
    & "And K.nConscod =2015  order by rtrim(ltrim(p.cpersnombre)) "
rsConv.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsConv.ActiveConnection = Nothing
Set GetServConvenios2 = rsConv
Set rsConv = Nothing
End Function

Public Function GetServConvenios(Optional ByVal nTipoCOnv As Integer) As ADODB.Recordset
Dim rsConv As ADODB.Recordset
Set rsConv = New ADODB.Recordset
rsConv.CursorLocation = adUseClient
sSql = "Select distinct Codigo=C.cPersCod, Institucion=ltrim(rtrim(P.cPersNombre)), cpersid=isnull(( Select top 1 cpersidnro from persid g where cpersidtpo in ('1','2') and g.cperscod=p.cperscod ),'')" _
    & "  From CaptacConvenio C INNER JOIN " & sDBPersona & "Persona P " _
    & "ON C.cPersCod = P.cPersCod INNER JOIN " & sDBComunes & "Constante K ON C.nConvCod = K.nConsValor " _
    & "And K.nConsValor = " & nTipoCOnv & "" _
    & " order by ltrim(rtrim(P.cPersNombre))  "
    
rsConv.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsConv.ActiveConnection = Nothing
Set GetServConvenios = rsConv
Set rsConv = Nothing
End Function

Public Function GetServConveniosArbol() As ADODB.Recordset
Dim rsConv As ADODB.Recordset
Set rsConv = New ADODB.Recordset
rsConv.CursorLocation = adUseClient
sSql = "Select C.cPersCod, P.cPersNombre, nNivel = 1 " _
    & "From CaptacConvenio C INNER JOIN " & sDBPersona & "Persona P " _
    & "ON C.cPersCod = P.cPersCod "
rsConv.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsConv.ActiveConnection = Nothing
Set GetServConveniosArbol = rsConv
Set rsConv = Nothing
End Function

Public Function GetServConveniosCodArbol() As ADODB.Recordset
Dim rsConv As ADODB.Recordset
Set rsConv = New ADODB.Recordset
rsConv.CursorLocation = adUseClient
sSql = "Select C.nConvCod, P.cPersNombre, nNivel = 1 " _
    & "From CaptacConvenio C INNER JOIN " & sDBPersona & "Persona P " _
    & "ON C.cPersCod = P.cPersCod "
rsConv.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsConv.ActiveConnection = Nothing
Set GetServConveniosCodArbol = rsConv
Set rsConv = Nothing
End Function

Public Sub EliminaServConvenio(ByVal sPersona As String)
sSql = "Delete CaptacConvenio Where cPersCod = '" & sPersona & "'"
dbCmact.Execute sSql
End Sub

Public Sub EliminaServConvCuentas(ByVal sPersona As String)
sSql = "Delete CaptacConvenioCta Where cPersCod = '" & sPersona & "'"
dbCmact.Execute sSql
End Sub
Public Sub ActualizaServConvenio(ByVal sPersona As String, ByVal nConvenio As CaptacConvenios)
sSql = " Update CaptacConvenio  set nConvCod=" & nConvenio _
    & " where cperscod='" & sPersona & "'"
dbCmact.Execute sSql
End Sub
Public Sub AgregaServConvenio(ByVal sPersona As String, ByVal nConvenio As CaptacConvenios)
sSql = "Insert CaptacConvenio (cPersCod,nConvCod) " _
    & "VALUES ('" & sPersona & "'," & nConvenio & ")"
dbCmact.Execute sSql
End Sub

Public Sub AgregaServConvenioCuenta(ByVal sPersona As String, ByVal sCuenta As String, _
            ByVal nTipoCuenta As CaptacConvTipoCuenta, nMontoMoraDia As Currency)
sSql = "Insert CaptacConvenioCta (cPersCod,cCtaCod,nTpoCuenta,nMontoMoraDia) " _
    & "VALUES ('" & sPersona & "','" & sCuenta & "'," & nTipoCuenta & "," & nMontoMoraDia & " )"
dbCmact.Execute sSql
End Sub

Public Function GetServPlanPagos(Optional sPersona As String = "", Optional nConvenio As CaptacConvenios) As ADODB.Recordset
Dim rsPlan As ADODB.Recordset

If sPersona <> "" Then
    sSql = "Select CONVERT(Varchar(10),dFecVenc,103) cVencimiento, nMontoCuota, nMontoGasto, " _
        & "bMora = CASE WHEN bAfectoMora = 1 THEN 'SI' ELSE 'NO' END, bFeriado = CASE WHEN bAfectoFeriado = 1 THEN 'SI' ELSE 'NO' END " _
        & "From CaptacConvenioPlanPag WHERE cPersCod = '" & sPersona & "'"
Else
    sSql = "Select PP.dFecVenc, PP.nMontoCuota, PP.nMontoGasto, PP.bAfectoMora, PP.bAfectoFeriado, PP.nNroCuota " _
        & "From CaptacConvenio C INNER JOIN CaptacConvenioPlanPag PP ON C.cPersCod = PP.cPersCod " _
        & "WHERE C.nConvCod = " & nConvenio & " ORDER BY PP.dFecVenc"
End If
Set rsPlan = New ADODB.Recordset
rsPlan.CursorLocation = adUseClient
rsPlan.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsPlan.ActiveConnection = Nothing
Set GetServPlanPagos = rsPlan
Set rsPlan = Nothing
End Function

Public Function GetServConvCuentas(Optional sPersona As String = "", Optional nConvenio As CaptacConvenios, Optional cSoloActivas As String = "N") As ADODB.Recordset
Dim rsPlan As ADODB.Recordset

If sPersona <> "" Then
  If cSoloActivas = "N" Then
    sSql = "Select [Nro Cuenta]=C.cCtaCod,Moneda=case when substring(c.cctacod,9,1)='1' then 'SOLES' else 'DOLARES' end ,[Fecha Apertura]=convert(char(10),ca.dapertura,103)  " _
        & "FROM CaptacConvenioCta C INNER JOIN " & sDBComunes & "Constante K ON C.nTpoCuenta = K.nConsValor " _
        & " JOIN CAPTACIONES CA ON CA.CCTACOD=C.CCTACOD " _
        & "WHERE    C.cPersCod = '" & sPersona & "' And K.nConsCod = " & gCaptacConvTipoCuenta & " ORDER BY " _
        & "C.nTpoCuenta"
  Else
    sSql = "Select [Nro Cuenta]=C.cCtaCod,Moneda=case when substring(c.cctacod,9,1)='1' then 'SOLES' else 'DOLARES' end ,[Fecha Apertura]=convert(char(10),ca.dapertura,103)  " _
        & "FROM CaptacConvenioCta C INNER JOIN " & sDBComunes & "Constante K ON C.nTpoCuenta = K.nConsValor " _
        & " JOIN CAPTACIONES CA ON CA.CCTACOD=C.CCTACOD " _
        & " join producto p on p.cctacod=ca.cctacod " _
        & "WHERE p.nprdestado not in (1300,1400) and C.cPersCod = '" & sPersona & "' And K.nConsCod = " & gCaptacConvTipoCuenta & " ORDER BY " _
        & "C.nTpoCuenta"
        
  End If
Else
    sSql = "Select CC.cCtaCod, K.cConsDescripcion, CC.nTpoCuenta, CC.nMontoMoraDia " _
        & "FROM CaptacConvenioCta CC INNER JOIN CaptacConvenio C ON CC.cPersCod = C.cPersCod INNER JOIN " _
        & sDBComunes & "Constante K ON CC.nTpoCuenta = K.nConsValor WHERE C.nConvCod = " & nConvenio & " And " _
        & "K.nConsCod = " & gCaptacConvTipoCuenta & " ORDER BY CC.nTpoCuenta"
End If
Set rsPlan = New ADODB.Recordset
rsPlan.CursorLocation = adUseClient
rsPlan.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsPlan.ActiveConnection = Nothing
Set GetServConvCuentas = rsPlan
Set rsPlan = Nothing
End Function

Public Sub EliminaPlanPagos(ByVal sPersona As String)
sSql = "Delete CaptacConvenioPlanPag Where cPersCod = '" & sPersona & "'"
dbCmact.Execute sSql
End Sub

Public Sub AgregaPlanPagos(ByVal sPersona As String, ByVal nNroCuota As Long, _
        ByVal dVencimiento As Date, ByVal nMontoCuota As Double, ByVal nMontoGasto As Double, _
        ByVal bMora As Boolean, ByVal bFeriado As Boolean)
sSql = "INSERT CaptacConvenioPlanPag (cPersCod,nNroCuota,dFecVenc,nMontoCuota,nMontoGasto,bAfectoMora,bAfectoFeriado) " _
    & "VALUES ( '" & sPersona & "'," & nNroCuota & ",'" & Format$(dVencimiento, "mm/dd/yyyy") & "'," & nMontoCuota & "," & nMontoGasto & "," & IIf(bMora, 1, 0) & "," & IIf(bFeriado, 1, 0) & ")"
dbCmact.Execute sSql
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
Set dbCmact = New ADODB.Connection
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
sSql = "Select C.nConceptoCod, C.cDescripcion Concepto, Monto = CASE CC.bPredefinido WHEN 1 THEN CC.nMonto WHEN 0 THEN 0 END FROM " _
    & "CapUNTConcepto C INNER JOIN CapUNTCurriculoConcepto CC INNER JOIN CapUNTMatricula M ON CC.cMatriculaCod = M.cMatriculaCod ON " _
    & "C.nConceptoCod = CC.nConceptoCod WHERE CC.cCurriculoTpo = '" & sCurriculo & "' " _
    & "AND CC.cMatriculaCod = '" & sTipoMatricula & "' And C.cEstado = '1' And C.nConceptoCod IN (" & gUNTMatReg & "," & gUNTMatExt & "," _
    & gUNTMatReg2daProf & "," & gUNTMatExt2daProf & ") And cSegProfesion = '" & s2daProf & "' And cIngresante = '" & sIngresante & "'"
Set rsMat = New ADODB.Recordset
rsMat.CursorLocation = adUseClient

rsMat.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMat.ActiveConnection = Nothing
Set GetUNTConceptosAlumno = rsMat
Set rsMat = Nothing
End Function

Public Function GetUNTAlumno(ByVal sMatricula As String) As ADODB.Recordset
Dim rsAlu As ADODB.Recordset
sSql = "Select A.cNombre, E.cDescripcion, E.cEscuelaCod, E.cCurriculo from CapUNTAlumno A " _
    & "INNER JOIN CapUNTEscuela E ON A.cEscuelaCod = E.cEscuelaCod WHERE A.cMatricula = '" _
    & sMatricula & "'"
Set rsAlu = New ADODB.Recordset
rsAlu.CursorLocation = adUseClient
rsAlu.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsAlu.ActiveConnection = Nothing
Set GetUNTAlumno = rsAlu
Set rsAlu = Nothing
End Function

Public Function GetConvenioCuenta(ByVal nConvenio As CaptacConvenios) As ADODB.Recordset
Dim rsCta As ADODB.Recordset
sSql = "Select CC.cCtaCod, K.cConsDescripcion, CC.nTpoCuenta, CC.nMontoMoraDia From " _
    & "CaptacConvenioCta CC INNER JOIN CaptacConvenio C ON CC.cPersCod = C.cPersCod INNER JOIN " _
    & "Where C.nConvCod = " & nConvenio
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetConvenioCuenta = rsCta
Set rsCta = Nothing
End Function

Public Function GetConvenioCuenta2(ByVal scperscod As String, ByVal nConvenio As CaptacConvenios) As ADODB.Recordset
Dim rsCta As ADODB.Recordset
sSql = " Select CC.cCtaCod,cast( CC.nTpoCuenta as char(1)) + K.cConsDescripcion , CC.nMontoMoraDia, ESTADO=(case when p.nprdestado IN (1400,1300) then 'NO' ELSE 'SI' END) From " _
    & " CaptacConvenioCta CC INNER JOIN CaptacConvenio C ON CC.cPersCod = C.cPersCod INNER JOIN " _
    & " (select * from constante where nconscod=2016) K ON K.nconsvalor=cc.ntpocuenta  " _
    & " join producto p on p.cctacod=cc.cctacod " _
    & " Where C.nConvCod = " & nConvenio & " and cc.cperscod='" & scperscod & "'"
    
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetConvenioCuenta2 = rsCta
Set rsCta = Nothing
End Function


Public Function GetUNTCurriculoConcepto(ByVal sCurriculo As String, ByVal sTipoMatricula As String, _
        ByVal sIngresante As String, ByVal s2daProf As String) As ADODB.Recordset
Dim rsCurr As ADODB.Recordset
sSql = "Select C.nConceptoCod, C.cDescripcion Concepto, Monto = CASE CC.bPredefinido WHEN 1 THEN CC.nMonto WHEN 0 THEN 0 END FROM " _
    & "CapUNTConcepto C INNER JOIN CapUNTCurriculoConcepto CC INNER JOIN CapUNTMatricula M ON CC.cMatriculaCod = M.cMatriculaCod ON " _
    & "C.nConceptoCod = CC.nConceptoCod WHERE CC.cCurriculoTpo = '" & sCurriculo & "' " _
    & "And CC.cMatriculaCod = '" & sTipoMatricula & "' And C.cEstado = '1' " _
    & "And cSegProfesion = '" & s2daProf & "' And cIngresante = '" & sIngresante & "'"
Set rsCurr = New ADODB.Recordset
rsCurr.CursorLocation = adUseClient
rsCurr.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCurr.ActiveConnection = Nothing
Set GetUNTCurriculoConcepto = rsCurr
Set rsCurr = Nothing
End Function

Public Function GetUNTConceptosOtros() As ADODB.Recordset
Dim rsCurr As ADODB.Recordset
sSql = "Select C.nConceptoCod, C.cDescripcion Concepto, Monto = 0 FROM " _
    & "CapUNTConcepto C WHERE C.cEstado = '2'"
Set rsCurr = New ADODB.Recordset
rsCurr.CursorLocation = adUseClient
rsCurr.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCurr.ActiveConnection = Nothing
Set GetUNTConceptosOtros = rsCurr
Set rsCurr = Nothing
End Function

Public Function GetUNTAlumnoNombre(ByVal sAlumno As String) As ADODB.Recordset
Dim rsUNT As ADODB.Recordset
sSql = "SELECT A.cMatricula [Num Matr], A.cNombre Nombre, E.cDescripcion Escuela, E.cEscuelaCod, E.cCurriculo from CapUNTAlumno A " _
    & "INNER JOIN CapUNTEscuela E ON A.cEscuelaCod = E.cEscuelaCod WHERE A.cNombre LIKE (RTRIM('" & sAlumno & "') + '%') " _
    & "ORDER BY A.cNombre"
Set rsUNT = New ADODB.Recordset
rsUNT.CursorLocation = adUseClient
rsUNT.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsUNT.ActiveConnection = Nothing
Set GetUNTAlumnoNombre = rsUNT
Set rsUNT = Nothing
End Function

Public Sub AgregaUNTOrdenPago(ByVal nMovNro As Long, ByVal sOrden As String, ByVal sTipo As String, _
        ByVal dFecPago As Date, ByVal sAlumno As String, ByVal nTotal As Double, _
        ByVal sAgencia As String)
sSql = "INSERT CapUNTOrdenPago (nMovNro,cOrdenCod,cTipo,dFecPago,cMatricula,nMonto,nSaldo,cFlag,cAgencia) " _
    & "VALUES (" & nMovNro & ",'" & sOrden & "','" & sTipo & "','" & Format$(dFecPago, "mm/dd/yyyy") & "','" & sAlumno & "'," & Trim(nTotal) & ",0,'1','" & sAgencia & "')"
dbCmact.Execute sSql
End Sub

Public Sub AgregaUNTOrdenPagoDet(ByVal nMovNro As Long, ByVal nConcepto As CapServUNTConcepto, _
        ByVal nMonto As Double)
sSql = "INSERT CapUNTOrdenPagoDet (nMovNro,nConceptoCod,nMonto) " _
    & "VALUES (" & nMovNro & "," & nConcepto & "," & nMonto & ")"
dbCmact.Execute sSql
End Sub
                
Public Function GetConvenioAlumno(ByVal nConvenio As CaptacConvenios, ByVal sCodigo As String) As ADODB.Recordset
Dim rsServ As ADODB.Recordset
Set rsServ = New ADODB.Recordset
rsServ.CursorLocation = adUseClient
sSql = "Select CD.cCodigo, CD.cNombre, CD.cReferencia From CaptacConvenio C INNER JOIN CaptacConvenioDet CD " _
    & "ON C.cPersCod = CD.cPersCod WHERE C.nConvCod = " & nConvenio & " AND CD.cCodigo = '" & sCodigo & "'"
rsServ.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsServ.ActiveConnection = Nothing
Set GetConvenioAlumno = rsServ
Set rsServ = Nothing
End Function

Public Function GetGiroTarifario() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    sSql = "Select nMoneda, nMontoIni, nMontoFin, nMontoMas, nValor, nTipoTarifa from GiroTarifario Order by nMoneda, nMontoIni"
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set GetGiroTarifario = rs
    Set rs = Nothing
End Function

Public Function GetGiroDatos(ByVal sCuenta As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
'    sSql = "Select PD.cCtaCod, PD.dPrdEstado, LTRIM(RTRIM(K.cConsDescripcion)) cTipo, PD.nSaldo, P.cPersNombre cRemitente, P.cPersDireccDomicilio, " _
'        & "D.cDestinatario, LTRIM(RTRIM(A.cAgeDescripcion)) cAgencia, ISNULL(ID.cPersIDnro,'') cDocID, D.cReferencia, D.cCodDest, D.cDireccion, P.cPersCod, " _
'        & "ISNULL(P.dPersNacCreac,'') dFecNac From " & sDBPersona & "PersID ID RIGHT JOIN " & sDBPersona & "Persona P INNER JOIN ProductoPersona PG " _
'        & "INNER JOIN Producto PD INNER JOIN Giro G INNER JOIN " _
'        & "(Select PD.cCtaCod, cDestinatario = Convert(Varchar(120),P.cPersNombre), cDireccion = P.cPersDireccDomicilio, " _
'        & "cReferencia = Convert(Varchar(255),ISNULL(ID.cPersIDnro,'')), cCodDest = P.cPersCod From " _
'        & sDBPersona & "PersID ID RIGHT JOIN " & sDBPersona & "Persona P INNER JOIN ProductoPersona PG INNER JOIN Producto PD ON " _
'        & "PG.cCtaCod = PD.cCtaCod ON P.cPersCod = PG.cPersCod ON ID.cPersCod = P.cPersCod Where PD.nPrdEstado NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ") " _
'        & "And PG.nPrdPersRelac = " & gGiroRelPersDestinatario & " " _
'        & "UNION " _
'        & "Select PD.cCtaCod, D.cDestinatario, cDireccion = '', D.cReferencia, cCodDest = '' From GiroDestinatario D INNER JOIN Producto PD ON " _
'        & "D.cCtaCod = PD.cCtaCod Where PD.nPrdEstado NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ")) D ON G.cCtaCod = D.cCtaCod ON " _
'        & "PD.cCtaCod = G.cCtaCod ON PG.cCtaCod = PD.cCtaCod ON P.cPersCod = PG.cPersCod ON ID.cPersCod = P.cPersCod INNER JOIN " & sDBComunes & "Agencias A ON " _
'        & "G.cAgenciaDest = A.cAgeCod INNER JOIN " & sDBComunes & "Constante K ON G.nPrdCtaTpo = K.nConsValor " _
'        & "Where PD.nPrdEstado NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ") And PG.nPrdPersRelac = " & gGiroRelPersRemitente & " And " _
'        & "K.nConsCod = " & gProductoCuentaTipo & " And G.cCtaCod = '" & sCuenta & "' Order by cDocID"

    'WIOR 20121107 Comentado y Agregado en Procedimiento
    sSql = "EXEC stp_sel_ObtenerDatosGiros '" & sCuenta & "'"
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set GetGiroDatos = rs
    Set rs = Nothing
End Function

Public Function GetGiroPendiente(ByVal sCodage As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient

sSql = "Select PD.cCtaCod, PD.dPrdEstado, LTRIM(RTRIM(K.cConsDescripcion)) cTipo, PD.nSaldo, " _
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
    & "Where PD.nPrdEstado NOT IN (1300,1400) And PG.nPrdPersRelac = " & gGiroRelPersRemitente & " And K.nConsCod = " & gProductoCuentaTipo _
    & " Order by PD.dPrdEstado"
    
    'By Capi 18082008 Se comento agencia para que se visualice todos los giros
    '& "Where PD.nPrdEstado NOT IN (1300,1400) And PG.nPrdPersRelac = " & gGiroRelPersRemitente & " And K.nConsCod = " & gProductoCuentaTipo & "  and   G.cagenciadest='" & sCodage & "'" _

rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetGiroPendiente = rs
Set rs = Nothing
End Function

Public Function GetGiroPendientexDoc(ByVal sCodage As String, ByVal pNumDoc As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
         
    'RECO20140423 ERS008-2014***********************************
    sSql = "exec stp_sel_GirosxAgenciayDoc '" & sCodage & "','" & pNumDoc & "'"
    'sSql = "exec stp_sel_GirosxAgenciayDoc '" & pNumDoc & "'"
    'RECO FIN***************************************************
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GetGiroPendientexDoc = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function GetGiroClaveSeguridad(ByVal sCuenta As String, Optional pbEsBanca As Boolean = False) As String 'add pbEsBanca
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sSql = "Select IsNULL(cClaveSeg,'') as cClaveSeg From Giro Where cCtaCod = '" & Trim(sCuenta) & "'"
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not rs.EOF And Not rs.BOF Then
        If rs.Fields("cClaveSeg") <> "" Then
            Dim objEnc As New COMConecta.DCOMClasIni
            'GetGiroClaveSeguridad = objEnc.Encripta(rs.Fields("cClaveSeg"), False)
            'ANDE 20190620: Modificación para obtener clave de giro para banca por internet
            If Not pbEsBanca Then
                GetGiroClaveSeguridad = objEnc.Encripta(rs.Fields("cClaveSeg"), False)
            Else
                GetGiroClaveSeguridad = rs.Fields("cClaveSeg")
            End If
            'END ANDE 20190620
            Set objEnc = Nothing
        Else
            GetGiroClaveSeguridad = ""
        End If
    Else
        GetGiroClaveSeguridad = ""
    End If
    Exit Function
End Function

Public Function GetPersonasExoLavDinero(Optional ByVal psPersCod As String = "") As ADODB.Recordset 'JACA 20110609 SE INCLUYO LA VARIABLE OPCIONAL
Dim rsLav As ADODB.Recordset
'MODIFICADO BY JACA 20110609
    'sSql = "Select L.cPersCod, RTRIM(P.cPersNombre) Nombre, (K.cConsDescripcion + SPACE(50) + Convert(Varchar(2),L.nEstado)) Estado, " _
    '    & "RTRIM(L.cComentario) Comentario From Persona P INNER JOIN PersExoLavDinero L ON P.cPersCod = L.cPersCod INNER JOIN Constante K ON " _
    '    & "L.nEstado = K.nConsValor Where L.cMovNro IN (Select MAX(L1.cMovNro) From PersExoLavDinero L1 Where L1.cPersCod = L.cPersCod) " _
    '    & "And K.nConsCOd = " & gPersEstLavDinero
    sSql = "exec stp_sel_PersonasExoLavDinero '" & psPersCod & "'"
'JACA END

Set rsLav = New ADODB.Recordset
rsLav.CursorLocation = adUseClient
rsLav.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsLav.ActiveConnection = Nothing
Set GetPersonasExoLavDinero = rsLav
Set rsLav = Nothing
End Function
'By Capi 28012008
Public Function GetPersonasDudLavDinero() As ADODB.Recordset
Dim rsLav As ADODB.Recordset
sSql = "Select L.cPersCod, RTRIM(P.cPersNombre) Nombre, (K.cConsDescripcion + SPACE(50) + Convert(Varchar(2),L.nEstado)) Estado, " _
    & "RTRIM(L.cDocumento01) Documento01,RTRIM(L.cDocumento02) Documento02,RTRIM(L.cComentario) Comentario From Persona P INNER JOIN PersDudLavDinero L ON P.cPersCod = L.cPersCod INNER JOIN Constante K ON " _
    & "L.nEstado = K.nConsValor Where L.cMovNro IN (Select MAX(L1.cMovNro) From PersDudLavDinero L1 Where L1.cPersCod = L.cPersCod) " _
    & "And K.nConsCOd = " & gPersCondControl

Set rsLav = New ADODB.Recordset
rsLav.CursorLocation = adUseClient
rsLav.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsLav.ActiveConnection = Nothing
Set GetPersonasDudLavDinero = rsLav
Set rsLav = Nothing
End Function
Public Function GetPersonaExoneradaLavDinero(ByVal sPersona As String) As ADODB.Recordset
    Dim rsLav As ADODB.Recordset
    'MODIFICADO BY JACA 20110609
        'sSql = "Select L.cPersCod, RTRIM(P.cPersNombre) Nombre, (K.cConsDescripcion + SPACE(50) + Convert(Varchar(2),L.nEstado)) Estado, L.nEstado, " _
        '    & "RTRIM(L.cComentario) Comentario From Persona P INNER JOIN PersExoLavDinero L ON P.cPersCod = L.cPersCod INNER JOIN Constante K ON " _
        '    & "L.nEstado = K.nConsValor Where L.cMovNro IN (Select MAX(L1.cMovNro) From PersDudLavDinero L1 Where L1.cPersCod = L.cPersCod) " _
        '    & "And K.nConsCOd = " & gPersCondControl & " And L.cPersCod = '" & sPersona & "'"
        sSql = "exec stp_sel_PersonaExoneradaLavDinero'" & sPersona & "'"
    'JACA END
    Set rsLav = New ADODB.Recordset
    rsLav.CursorLocation = adUseClient
    rsLav.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsLav.ActiveConnection = Nothing
    Set GetPersonaExoneradaLavDinero = rsLav
    Set rsLav = Nothing
End Function

Public Function GetInfoRP(ByVal sAgencia As String) As ADODB.Recordset
Dim rsrp As ADODB.Recordset

sSql = " select R.CUSER , P.CPERSNOMBRE , R.CPERSCOD "
sSql = sSql & " from rrhh R "
sSql = sSql & " JOIN PERSONA P ON  P.CPERSCOD=R.CPERSCOD "
sSql = sSql & " where nrhestado=201 and crhcod like 'E%' "
sSql = sSql & " and careacod in ('026','066') and cagenciaactual='" & sAgencia & "'"
sSql = sSql & " ORDER BY P.CPERSNOMBRE "

Set rsrp = New ADODB.Recordset
rsrp.CursorLocation = adUseClient
rsrp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsrp.ActiveConnection = Nothing
Set GetInfoRP = rsrp
Set rsrp = Nothing
End Function


Public Function GetInfoActa(ByVal sCodUsu As String) As ADODB.Recordset
Dim rsrp As ADODB.Recordset

sSql = " select FECHA=CONVERT(CHAR(10),FechaIncautacion,103),PORTADOR=p.CPERSNOMBRE , NroActa ,FechaReg=CAST(LEFT(CMOVNRO,8) AS DATETIME), CodPortador "
sSql = sSql & " from actasincautacion a "
sSql = sSql & " join persona p on p.cperscod= a.CodPortador "
sSql = sSql & " where right(cmovnro,4)='" & sCodUsu & "'"
sSql = sSql & " ORDER BY NROACTA DESC "

Set rsrp = New ADODB.Recordset
rsrp.CursorLocation = adUseClient
rsrp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsrp.ActiveConnection = Nothing
Set GetInfoActa = rsrp
Set rsrp = Nothing
End Function

Public Function GetInfoActaCompleta(ByVal sCodUsu As String, ByVal sNroActa As String) As ADODB.Recordset
Dim rsrp As ADODB.Recordset

sSql = " Select FECHA=FechaIncautacion,PORTADOR=p.CPERSNOMBRE,DNI=(SELECT TOP 1 CPERSIDNRO "
sSql = sSql & " FROM PERSID I WHERE I.CPERSCOD=A.CodPortador AND I.CPERSIDTPO=1) , VERIFICADOR=r.cuser+': '+P1.cpersnombre , NroActa,CodPortador,Codusuverificador,a.cmoneda,FechaReg=cast(left(a.cmovnro,8) as datetime) "
sSql = sSql & " from actasincautacion a "
sSql = sSql & " join persona p on p.cperscod= a.CodPortador "
sSql = sSql & " join persona p1 on p1.cperscod= a.codusuverificador "
sSql = sSql & " join rrhh r on r.cperscod= a.codusuverificador "
sSql = sSql & " where codusudetector='" & sCodUsu & "' and nroacta='" & sNroActa & "'"
sSql = sSql & " ORDER BY cast(NROACTA as bigint) DESC "

Set rsrp = New ADODB.Recordset
rsrp.CursorLocation = adUseClient
rsrp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsrp.ActiveConnection = Nothing
Set GetInfoActaCompleta = rsrp
Set rsrp = Nothing
End Function

Public Function GetInfoDetActa(ByVal sNroActa As String) As ADODB.Recordset
Dim rsrp As ADODB.Recordset

sSql = " select serienum,fechaemision=CONVERT(CHAR(10),FECHAEMISION,103),denominacion "
sSql = sSql & " From detactaincautacion "
sSql = sSql & " Where nroacta = '" & sNroActa & "' order by nitem"

Set rsrp = New ADODB.Recordset
rsrp.CursorLocation = adUseClient
rsrp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsrp.ActiveConnection = Nothing
Set GetInfoDetActa = rsrp
Set rsrp = Nothing
End Function

Public Function GrabaActa(ByVal CodPortador As String, ByVal FechaIncautacion As Date, ByVal CodUsuDetector As String, ByVal COdUsuVerificador As String, ByVal cMovNro As String, ByVal cMoneda As String, ByVal RSTMP As ADODB.Recordset, ByRef NroActa As String) As Boolean
Dim rsrp As ADODB.Recordset, sSql As String
Dim sNroActa As String, i As Integer

i = 0
On Error GoTo ERRMEN
dbCmact.BeginTrans

sSql = " select NUEVONUMERO=isnull(MAX(NROACTA),0)+1 "
sSql = sSql & " From actaSincautacion "
'ssql = ssql & " Where nroacta = '" & sNroActa & "'"
Set rsrp = New ADODB.Recordset
rsrp.CursorLocation = adUseClient
rsrp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
If Not rsrp.EOF Then

    'sNroActa = String(10 - Len(rsrp!NuevoNumero), "0") & rsrp!NuevoNumero
    
    sNroActa = rsrp!NuevoNumero
    sSql = " Insert Into  ActasIncautacion "
    sSql = sSql & " values('" & sNroActa & "','" & CodPortador & "','" & Format(FechaIncautacion, "yyyy-mm-dd") & "','" & CodUsuDetector & "','" & COdUsuVerificador & "','" & cMovNro & "','" & cMoneda & "') "
    dbCmact.Execute sSql
        
    While Not RSTMP.EOF
        i = i + 1
        sSql = " Insert Into  DetActaIncautacion "
        sSql = sSql & " values('" & sNroActa & "'," & i & ",'" & RSTMP.Fields("SERIE Y NUMERACION").Value & "', '" & Format(Left(RSTMP.Fields(1).Value, 10), "mm/dd/yyyy") & "'," & CCur(RSTMP.Fields(2).Value) & ") "
        dbCmact.Execute sSql
            
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

Err.Raise Err.Number, "No se guardó la información. Error:" & Err.Number & vbCrLf & "Porfavor, Inténtelo otra vez", Err.Description

dbCmact.RollbackTrans

End Function

Public Function GetPersonaHistExoLavDinero(ByVal sPerscod As String) As ADODB.Recordset
Dim rsLav As ADODB.Recordset
'MODIFICADO BY JACA 20110609
    'sSql = "Select (SUBSTRING(L.cMovNro,7,2) + '/' + SUBSTRING(L.cMovNro,5,2) + '/' + SUBSTRING(L.cMovNro,1,4)) Fecha, " _
    '    & "(K.cConsDescripcion + SPACE(50) + Convert(Varchar(2),L.nEstado)) Estado, " _
    '    & "RTRIM(L.cComentario) Comentario, L.cPersCod, L.cMovNro From PersExoLavDinero L INNER JOIN Constante K ON " _
    '    & "L.nEstado = K.nConsValor Where K.nConsCOd = " & gPersEstLavDinero & " And L.cPersCod = '" & sPersCod & "' " _
    '    & "Order by L.cMovNro"
    sSql = "exec stp_sel_PersonaHistExoLavDinero '" & sPerscod & "'"
'JACA END
Set rsLav = New ADODB.Recordset
rsLav.CursorLocation = adUseClient
rsLav.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsLav.ActiveConnection = Nothing
Set GetPersonaHistExoLavDinero = rsLav
Set rsLav = Nothing
End Function
'By Capi 28012008
Public Function GetPersonaHistDudLavDinero(ByVal sPerscod As String) As ADODB.Recordset
Dim rsLav As ADODB.Recordset
sSql = "Select (SUBSTRING(L.cMovNro,7,2) + '/' + SUBSTRING(L.cMovNro,5,2) + '/' + SUBSTRING(L.cMovNro,1,4)) Fecha, " _
    & "(K.cConsDescripcion + SPACE(50) + Convert(Varchar(2),L.nEstado)) Estado, " _
    & "RTRIM(L.cComentario) Comentario, L.cPersCod, L.cMovNro From PersExoLavDinero L INNER JOIN Constante K ON " _
    & "L.nEstado = K.nConsValor Where K.nConsCOd = " & gPersEstLavDinero & " And L.cPersCod = '" & sPerscod & "' " _
    & "Order by L.cMovNro"

Set rsLav = New ADODB.Recordset
rsLav.CursorLocation = adUseClient
rsLav.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsLav.ActiveConnection = Nothing
Set GetPersonaHistDudLavDinero = rsLav
Set rsLav = Nothing
End Function
Public Function GetTitularExoLavDinero(ByVal sCuenta As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
'MODIFICADO BY JACA 20110609
    'sSql = "Select PP.cPersCod, L.cPersCod PersonaExo From ProductoPersona PP LEFT JOIN " _
    '    & "(Select L.cPersCod, L.nEstado From PersExoLavDinero L Where cMovNro IN (Select MAX(cMovNro) From " _
    '    & "PersExoLavDinero L1 Where L1.cPersCod = L.cPersCod And L1.nEstado = " & gPersLavDinSinControl & ") " _
    '    & "And L.nEstado = " & gPersLavDinSinControl & ") L ON PP.cPersCod = L.cPersCod Where " _
    '    & "PP.nPrdPersRelac = " & gCapRelPersTitular & " And PP.cCtaCod = '" & sCuenta & "'"
    sSql = "exec stp_sel_TitularExoLavDinero'" & sCuenta & "'"
'JACA END

rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetTitularExoLavDinero = rs
Set rs = Nothing
End Function

Public Sub AgregaPersExoLavDinero(ByVal sPerscod As String, ByVal sMovNro As String, _
    ByVal nEstado As PersEstLavDinero, ByVal sComentario As String)
sSql = "Insert PersExoLavDinero (cPersCod,cMovNro,nEstado,cComentario) " _
    & "Values ('" & sPerscod & "','" & sMovNro & "'," & nEstado & ",'" & sComentario & "')"
dbCmact.Execute sSql
End Sub


'By Capi 28012008
Public Sub AgregaPersDudLavDinero(ByVal sPerscod As String, ByVal sMovNro As String, _
    ByVal nEstado As PersEstLavDinero, ByVal sComentario As String, sDocumento01 As String, _
    sDocumento02 As String)
    
sSql = "Insert PersDudLavDinero (cPersCod,cMovNro,nEstado,cComentario,cDocumento01,cDocumento02) " _
    & "Values ('" & sPerscod & "','" & sMovNro & "'," & nEstado & ",'" & sComentario & "','" & sDocumento01 & "','" & sDocumento02 & "')"
dbCmact.Execute sSql
End Sub
Public Sub ActualizaPersExoLavDinero(ByVal sPerscod As String, ByVal sMovNro As String, _
    ByVal sComentario As String, ByVal sMovNroAnt As String)
sSql = "Update PersExoLavDinero Set cComentario = '" & sComentario & "', cMovNro = '" & sMovNro & "' " _
    & "Where cPersCod = '" & sPerscod & "' And cMovNro = '" & sMovNroAnt & "'"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaChequeFechaValorizacion(ByVal sNroDoc As String, ByVal sPerscod As String, _
        ByVal sIFTpo As String, ByVal dNuevaFecha As Date)
sSql = "Update DocRec Set dValorizaRef = '" & Format$(dNuevaFecha, "mm/dd/yyyy") & "' " _
    & "Where nTpoDoc = " & TpoDocCheque & " And cNroDoc = '" & sNroDoc & "' And cPersCod = '" & sPerscod & "' " _
    & "And cIFTpo = '" & sIFTpo & "'"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaSaldoAnteriorAho(ByVal sCuenta As String, ByVal nSaldoAnt As Double)
sSql = "Update CaptacAhorros Set nSaldoAnterior = " & nSaldoAnt & "  WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaAbonoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)

If bActExtracto Then
    sSql = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSql = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If
dbCmact.Execute sSql
sSql = "Update Captaciones Set nSaldoDisp = nSaldoDisp + " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
    & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaCargoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)

If bActExtracto Then
    sSql = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSql = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If
dbCmact.Execute sSql
sSql = "Update Captaciones Set nSaldoDisp = nSaldoDisp - " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
    & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub AgregaDocRecEstado(ByVal sNroDoc As String, sPerscod As String, _
        ByVal sTipoIF As String, ByVal nEstado As ChequeEstado, ByVal sMovNro As String, Optional psCtaCheque As String = "")
        
sSql = "INSERT INTO DocRecEst(nTpoDoc,cNroDoc,cPersCod,cIFTpo,nEstado,cMovNro,CIFcta) " _
     & "VALUES (" & TpoDocCheque & ", '" & sNroDoc & "','" & sPerscod & "','" & Format(sTipoIF, "00") & "'," _
     & nEstado & ",'" & sMovNro & "','" & psCtaCheque & "')"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaEstadoCheque(ByVal sNroDoc As String, sPerscod As String, _
        ByVal sTipoIF As String, ByVal nEstado As ChequeEstado, Optional dFechaVal As Date)
        
If dFechaVal <> CDate("12:00:00 AM") Then
    sSql = "Update DocRec Set nEstado = " & nEstado & ", dValorizacion = '" & Format$(dFechaVal, "mm/dd/yyyy") & "' " _
        & "Where nTpoDoc = " & TpoDocCheque & " And cNroDoc = '" & sNroDoc & "' And cPersCod = '" & sPerscod & "' " _
        & "And cIFTpo = '" & sTipoIF & "'"
Else
'    sSQL = "Update DocRec Set nEstado = " & nEstado & " " _
'        & "Where nTpoDoc = " & TpoDocCheque & " And cNroDoc = '" & sNroDoc & "' And cPersCod = '" & sPersCod & "' " _
'        & "And cIFTpo = '" & sTipoIF & "'"
    sSql = "DELETE DocRec " _
        & "Where nTpoDoc = " & TpoDocCheque & " And cNroDoc = '" & sNroDoc & "' And cPersCod = '" & sPerscod & "' " _
        & "And cIFTpo = '" & sTipoIF & "'"

End If
dbCmact.Execute sSql
End Sub

Public Sub ActualizaFechaRenovacionPF(ByVal sCuenta As String, ByVal dAperturaChq As Date)
sSql = "Update CaptacPlazoFijo Set dRenovacion = '" & Format$(dAperturaChq, "mm/dd/yyyy") & "' " _
    & "Where cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaSaldoDisponiblePF(ByVal sCuenta As String, ByVal dAperturaChq As Date, ByVal nMonto As Double)
sSql = "Update Captaciones Set nSaldoDisp = " & nMonto & ", dUltCierre = '" & Format$(DateAdd("d", -1, dAperturaChq), "mm/dd/yyyy") & "' " _
    & "Where cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub
        
Public Sub ActualizaSaldoRetiroCTS(ByVal sCuenta As String, ByVal nSaldoRet As Double, nIntSaldo As Double)
sSql = "Update CaptacCTS Set nSaldRetiro = " & nSaldoRet & ", nIntSaldo = " & nIntSaldo & " " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub AgregaMovCap(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nMonto As Double, _
        ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double)
sSql = "INSERT MovCap (nMovNro,cOpeCod,cCtaCod,nMonto,nSaldoDisponible,nSaldoContable) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nMonto & "," & nSaldoDisp & "," & nSaldoCnt & ")"
dbCmact.Execute sSql
End Sub

Public Sub AgregaMovCapDet(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nConcepto As CaptacConcepto, ByVal nMonto As Double)
        
sSql = "INSERT MovCapDet (nMovNro,cOpeCod,cCtaCod,nConceptoCod,nMonto) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nConcepto & "," & nMonto & ")"
dbCmact.Execute sSql
End Sub

Public Function GetPorcentajeRetiroCTS(ByVal sCuenta As String, ByVal sNroDoc As String, ByVal sPerscod As String, _
            ByVal sIFTpo As String) As Double
Dim rsChq As ADODB.Recordset
sSql = "Select P.nPorcentajeAbono From MovParamCTS P INNER JOIN Mov M INNER JOIN DocRecCapta C ON " _
    & "M.nMovNro = C.nMovNro ON P.nMovNro = M.nMovNro Where C.cNroDoc = '" & sNroDoc & "' And " _
    & "C.cPersCod = '" & sPerscod & "' And C.cIFTpo = '" & sIFTpo & "' And C.cCtaCod = '" & sCuenta & "'"

Set rsChq = New ADODB.Recordset
rsChq.CursorLocation = adUseClient
rsChq.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsChq.ActiveConnection = Nothing
GetPorcentajeRetiroCTS = rsChq("nPorcentajeAbono")
rsChq.Close
Set rsChq = Nothing
End Function

Public Function GetProvisionCTS(ByVal nMovNro As Long)
Dim rsChq As ADODB.Recordset
sSql = "Select MCD.nConceptoCod From Mov M JOIN MovCap MC JOIN MovCapDet MCD ON MC.nMovNro = MCD.nMovNro " _
    & "And MC.cCtaCod = MCD.cCtaCod And MC.cOpeCod = MCD.cOpeCod ON M.nMovNro = MC.nMovNro And " _
    & "M.cOpeCod = MC.cOpeCod Where M.nMovNro = " & nMovNro & " And MCD.nConcepto = " & gConcProvision & " "

Set rsChq = New ADODB.Recordset
rsChq.CursorLocation = adUseClient
rsChq.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsChq.ActiveConnection = Nothing
GetProvisionCTS = rsChq("nConceptoCod")
rsChq.Close
Set rsChq = Nothing
End Function

Public Function GetChequeCuenta(ByVal sNroDoc As String, ByVal sPerscod As String, ByVal sIFTpo As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSql = "Select cCtaCod, nMonto From DocRecCapta Where cNroDoc = '" & sNroDoc & "' And " _
    & "cPersCod = '" & sPerscod & "' And cIFTpo = '" & sIFTpo & "' Order by cCtaCod"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetChequeCuenta = rs
Set rs = Nothing
End Function

Public Function AgregaGiro(ByVal sCuenta As String, ByVal dApertura As Date, ByVal nMonto As Double, _
    ByVal nTipoGiro As ProductoCuentaTipo, ByVal sAgenciaDest As String, ByVal psClaveGiro As String)
    
    Dim sClaveGiro As String
    
    sSql = "Insert Producto (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc) " _
        & "Values ('" & sCuenta & "',0," & nMonto & "," & gCapEstActiva & ",'" & Format$(dApertura & " " & Time, "mm/dd/yyyy hh:mm:ss") & "',0)"
    dbCmact.Execute sSql
    
    If psClaveGiro <> "" Then
        Dim objEnc As New COMConecta.DCOMClasIni
        sClaveGiro = objEnc.Encripta(psClaveGiro)
        Set objEnc = Nothing
        sSql = "Insert Giro (cCtaCod,dVigencia,nPrdCtaTpo,cAgenciaDest,cClaveSeg,nClavePinPad) " _
            & "Values ('" & sCuenta & "','" & Format$(dApertura, "mm/dd/yyyy") & "'," & nTipoGiro & ",'" & sAgenciaDest & "','" & sClaveGiro & "', 1)"
    Else
        sSql = "Insert Giro (cCtaCod,dVigencia,nPrdCtaTpo,cAgenciaDest) " _
        & "Values ('" & sCuenta & "','" & Format$(dApertura, "mm/dd/yyyy") & "'," & nTipoGiro & ",'" & sAgenciaDest & "')"
    End If
    dbCmact.Execute sSql
End Function

Public Sub AgregaProdPersGiro(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As GiroRelacPersona)
    sSql = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " _
        & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ")"
    dbCmact.Execute sSql
End Sub
'Modificado x MADM 20101011 -- pNumDoc
Public Sub AgregaGiroDestinatario(ByVal sCuenta As String, ByVal sDestinatario As String, _
        ByVal sReferencia As String, Optional ByVal pNumDoc As String)
sSql = "Insert GiroDestinatario (cCtaCod,cDestinatario,cReferencia,cNumDoc) " _
    & "Values ('" & sCuenta & "','" & sDestinatario & "','" & sReferencia & "','" & pNumDoc & "')"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaEstadoProducto(ByVal sCuenta As String, ByVal nEstado As CaptacEstado, _
        ByVal dFecha As Date)
'Actualiza el ultimo estado a la tabla de producto
sSql = "Update Producto Set nPrdEstado = " & nEstado & ", " _
    & "dPrdEstado = '" & Format$(dFecha & " " & Time, "mm/dd/yyyy hh:mm:ss") & "' WHERE " _
    & "cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

' NUEVO - CMCPL
Public Sub AgregaServSATTributo(ByVal sCodTributo As String, ByVal sTributo As String, ByVal sAno As String, _
        ByVal sRecibo As String, ByVal sDigCheq As String, ByVal sPeriodo As String, ByVal nValDeuda As Double, _
        ByVal nValDerEmis As Double, ByVal nValAjuste As Double, ByVal nValIntMor As Double, ByVal dFecVenc As Date, _
        ByVal sNombre As String, ByVal sCuenta As String, ByVal sContrib As String, ByVal nValGastos As Double, _
        ByVal nValCostas As Double, ByVal sCodSis As String, ByVal nTotal As Double)

sSql = "INSERT CapServSATTributos (cCodTributo,cTributo,cAno,cRecibo,cDigCheq,cPeriodo,nValDeuda,nValDerEmis,nValAjuste, " _
    & "nValIntMor , dFecVenc, cNombre, cCuenta, cContrib, nValGastos, nValCostas, cCodSis, nTotal)  " _
    & "VALUES ('" & sCodTributo & "','" & sTributo & "','" & sAno & "','" & sRecibo & "','" & sDigCheq & "', " _
    & "'" & sPeriodo & "'," & nValDeuda & "," & nValDerEmis & "," & nValAjuste & ", " & nValIntMor & ", " _
    & "'" & Format$(dFecVenc, "mm/dd/yyyy") & "','" & sNombre & "','" & sCuenta & "','" & sContrib & "', " _
    & nValGastos & "," & nValCostas & ",'" & sCodSis & "'," & nTotal & ")"
dbCmact.Execute sSql
End Sub
' NUEVO - CMCPL
Public Sub AgregaServSATpapeletas(ByVal sCodTributo As String, ByVal sTributo As String, ByVal sAno As String, _
        ByVal sPlaca As String, ByVal sDigCheq As String, ByVal sPeriodo As String, ByVal nImporte As Double, _
        ByVal nValReincide As Double, ByVal nValGastAdm As Double, ByVal nValCostas As Double, ByVal dFecVenc As Date, _
        ByVal dFecInf As Date, ByVal dFecLect As String, ByVal dFecNoti As String, ByVal sPapeleta As String, _
        ByVal sTipo As String, ByVal sFalta As String, ByVal sCodAux As String, ByVal sLicCond As String)

sSql = "INSERT CapServSATPapeletas (cCodTributo,cTributo,cAno,cPlaca,cDigCheq,cPeriodo,nImporte,nValReincide,nValGastAdm, " _
    & "nValCostas , dFecVenc, dFecInf, dFecLect, dFecNoti, sPapeleta, sTipo, sFalta, sCodAux, sLicCond )  " _
    & "VALUES ('" & sCodTributo & "','" & sTributo & "','" & sAno & "','" & sPlaca & "','" & sDigCheq & "', " _
    & "'" & sPeriodo & "'," & nImporte & "," & nValReincide & "," & nValGastAdm & ", " & nValCostas & ", " _
    & "'" & Format$(dFecVenc, "mm/dd/yyyy") & "', '" & Format$(dFecInf, "mm/dd/yyyy") & "', " _
    & "'" & Format$(dFecLect, "mm/dd/yyyy") & "', '" & Format$(dFecNoti, "mm/dd/yyyy") & "', " _
    & "'" & sPapeleta & "','" & sTipo & "','" & sFalta & "','" & sCodAux & "','" & sLicCond & "')"
    
dbCmact.Execute sSql
End Sub
' NUEVO - CMCPL
Public Sub EliminaServSATTributo()
sSql = "TRUNCATE TABLE CapServSATTributos"
dbCmact.Execute sSql
End Sub
' NUEVO - CMCPL
Public Sub EliminaServSATPapeletas()
'sSql = "TRUNCATE TABLE CapServSATPapeletas"
'dbCmact.Execute sSql
End Sub
' NUEVO - CMCPL
Public Function GetServSATPapeletaLincencia(ByVal sLicencia As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSql = "Select cPlaca, nImporte, nValGastAdm, nValCostas, dFecInf, sPapeleta, sTipo, sFalta, " _
    & "sLicCond From CapServSATPapeletas Where sLicCond = '" & sLicencia & "' Order by dFecInf"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetServSATPapeletaLincencia = rs
Set rs = Nothing
End Function
' NUEVO - CMCPL
Public Function GetServSATPapeletaPlaca(ByVal sPlaca As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSql = "Select cPlaca, nImporte, nValGastAdm, nValCostas, dFecInf, sPapeleta, sTipo, sFalta, " _
    & "sLicCond From CapServSATPapeletas Where cPlaca = '" & sPlaca & "' Order by dFecInf"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetServSATPapeletaPlaca = rs
Set rs = Nothing
End Function
' NUEVO - CMCPL
Public Function GetServSATPapeletaNPapeleta(ByVal sPapeleta As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSql = "Select cPlaca, nImporte, nValGastAdm, nValCostas, dFecInf, sPapeleta, sTipo, sFalta, " _
    & "sLicCond From CapServSATPapeletas Where sPapeleta = '" & sPapeleta & "' Order by dFecInf"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetServSATPapeletaNPapeleta = rs
Set rs = Nothing
End Function
' NUEVO - CMCPL
Public Function GetServSATRecibo(ByVal sCuenta As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSql = "Select cCodtributo, cnombre, cTributo, cAno, cRecibo, nValDeuda, nValDerEmis, nValCostas, nValAjuste, " _
    & "nValIntMor,cCuenta,cContrib, nValGastos From CapServSATtributos Where cCuenta = '" & sCuenta & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetServSATRecibo = rs
Set rs = Nothing
End Function
' NUEVO - CMCPL
Public Function GetServSATDerecho(ByVal sCodigo As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSql = "Select cdistrito, ccodDerecho, cTipValor, cDescrip, nValDerecho From CapServicioDerecho Where ccodDerecho= '" & sCodigo & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetServSATDerecho = rs
Set rs = Nothing
End Function
' NUEVO - CMCPL CRSF 08/10
Public Function GetSatMontoFecha(ByVal dInicio As String, ByVal dFin As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSql = "Select a.nMovnro,substring(a.cMovNro,22,4) USUA, a.cOpeCod, b.nmonto, " _
    & " a.cMovdesc,substring(cmovNro,7,2)+'/'+substring(cMovNro,5,2)+'/'+substring(cMovNro,1,4) fecha" _
    & " From mov  a, movservicios b Where a.nmovnro = b.nmovnro  and  cOpeCod in ('300106','300105')  and " _
    & "substring(cmovNro,7,2)+'/'+substring(cMovNro,5,2)+'/'+substring(cMovNro,1,4) = '" & dInicio & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetSatMontoFecha = rs
Set rs = Nothing
End Function

' NUEVO - CMCPL CRSF 10/10
Public Function GetSatMontoFechaDis(ByVal dInicio As String, ByVal dFin As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSql = "Select a.cOpeCod,b.nPrdConcepto,sum(b.nMonto) nMonto From mov a, movserviciosdet b " _
        & " where a.nmovnro = b.nmovnro  and  a.cOpeCod in ('300106','300105')  and " _
    & "substring(cmovNro,7,2)+'/'+substring(cMovNro,5,2)+'/'+substring(cMovNro,1,4) = '" & dInicio & "'" _
    & " group by  a.cOpeCod,b.nPrdConcepto order by a.cOpeCod,b.nPrdConcepto"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetSatMontoFechaDis = rs
Set rs = Nothing
End Function
'NUEVO - CMCPL CRSF 11/06 - Monto de los Parametros de Servicios para Distrib
Public Function GetSatDistribucion(ByVal sParam As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSql = "Select cCtacodAbono,nComision from CapServicioParam " _
        & "Where nservcod = 300000 And nTipoComision = '" & sParam & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetSatDistribucion = rs
Set rs = Nothing
End Function

'NUEVO - CMCPL CRSF 11/06 - Suma de Los Totales
Public Function GetSatSuma(ByVal sParam As String) As Double
Dim rs As ADODB.Recordset
sSql = "Select sum(b.nMonto) nMonto from mov a, movserviciosdet b " _
        & "where a.nMovNro = b.nMovNro and cOpeCod in ('300105','300106') and nPrdConcepto ='" & sParam & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rs.ActiveConnection = Nothing
GetSatSuma = rs("nMonto")
rs.Close
Set rs = Nothing
End Function

Public Function GetNumSolcitudCapTasaEspecial() As Long
Dim rs As ADODB.Recordset
sSql = "Select nNumSolicitud = ISNULL(MAX(nNumSolicitud),0) + 1 From CapTasaEspecial"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
GetNumSolcitudCapTasaEspecial = rs("nNumSolicitud")
rs.Close
Set rs = Nothing
End Function

'Modificado RIRO 20130327 - Se adicionaron campos opcionales: nTasaTarifada, nTasaSolicitada y sSubProducto
Public Sub AgregaCapTasaEspecial(ByVal nNumSolicitud As Integer, ByVal sPersona As String, _
                                 ByVal nProd As Producto, ByVal nMon As Moneda, ByVal nEstado As Integer, _
                                 ByVal sMovNro As String, ByVal nTasa As Double, ByVal sComentario As String, _
                                 ByVal nMonto As Double, Optional ByVal sCuenta As String = "", _
                                 Optional ByVal nPlazo As Integer = 0, Optional ByVal nExtorno As Integer = 99, _
                                 Optional ByVal bPermanente As Boolean, Optional ByVal sSubProducto As Producto = 0, _
                                 Optional ByVal nTasaTarifada As Double, Optional ByVal nTasaSolicitada As Double)
                                 
Dim rs As ADODB.Recordset

If nExtorno <> 99 Then
        sSql = " Insert CapTasaEspecial (nNumSolicitud,cPersCod,nProducto,nMoneda,nEstado,cMovNro,nTasa,cComentario,cCtaCod,nMonto,nPlazo,nextorno,bPermanente, cSubProducto, nTasaTarif, nTasaSolicitada) " _
                 & "  VALUES (" & nNumSolicitud & ",'" & sPersona & "'," & Trim(nProd) & "," & Trim(nMon) & "," & Trim(nEstado) & ",'" & sMovNro & "'," & Trim(nTasa) & ",'" & sComentario & "','" & sCuenta & "'," & Trim(nMonto) & "," & Trim(nPlazo) & "," & nExtorno & "," & IIf(bPermanente, 1, 0) & ", " & sSubProducto & ", " & nTasaTarifada & ", " & nTasaSolicitada & " ) "

'    sSql = " Insert CapTasaEspecial (nNumSolicitud,cPersCod,nProducto,nMoneda,nEstado,cMovNro,nTasa,cComentario,cCtaCod,nMonto,nPlazo,nextorno,bPermanente) " _
'         & "  VALUES (" & nNumSolicitud & ",'" & sPersona & "'," & Trim(nProd) & "," & Trim(nMon) & "," & Trim(nEstado) & ",'" & sMovNro & "'," & Trim(nTasa) & ",'" & sComentario & "','" & sCuenta & "'," & Trim(nMonto) & "," & Trim(nPlazo) & "," & nExtorno & "," & IIf(bPermanente, 1, 0) & " ) "

Else
        sSql = " Insert CapTasaEspecial (nNumSolicitud,cPersCod,nProducto,nMoneda,nEstado,cMovNro,nTasa,cComentario,cCtaCod,nMonto,nPlazo ,bpermanente, cSubProducto, nTasaTarif, nTasaSolicitada) " _
                & " VALUES (" & nNumSolicitud & ",'" & sPersona & "'," & Trim(nProd) & "," & Trim(nMon) & "," & Trim(nEstado) & ",'" & sMovNro & "'," & Trim(nTasa) & ",'" & sComentario & "','" & sCuenta & "'," & Trim(nMonto) & "," & Trim(nPlazo) & "," & IIf(bPermanente, 1, 0) & ", " & sSubProducto & ", " & nTasaTarifada & ", " & nTasaSolicitada & ") "
'    sSql = " Insert CapTasaEspecial (nNumSolicitud,cPersCod,nProducto,nMoneda,nEstado,cMovNro,nTasa,cComentario,cCtaCod,nMonto,nPlazo,bpermanente) " _
'         & "  VALUES (" & nNumSolicitud & ",'" & sPersona & "'," & Trim(nProd) & "," & Trim(nMon) & "," & Trim(nEstado) & ",'" & sMovNro & "'," & Trim(nTasa) & ",'" & sComentario & "','" & sCuenta & "'," & Trim(nMonto) & "," & Trim(nPlazo) & "," & IIf(bPermanente, 1, 0) & ") "
End If
    
dbCmact.Execute sSql

End Sub

Public Function GetEstadoTPCta(ByVal sCuenta As String) As ADODB.Recordset
  Dim rs As ADODB.Recordset
  
  
  
  sSql = " Select  nNumSolicitud,cMovNro,cPersCod,nProducto,nMoneda,cCtaCod,nMonto,nPlazo,ntasa,nestado,cComentario,NEXTORNO,bpermanente=isnull(bpermanente,0) from captasaespecial where nestado=2 and  cctacod='" & sCuenta & "' "
  sSql = sSql & " and left(cmovnro,14)+cast(nnumsolicitud as varchar(12)) in (select max(left(c.cmovnro,14)+cast(c.nnumsolicitud as varchar(12))) from captasaespecial c where cctacod='" & sCuenta & "')"
     
  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseClient
  rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
  
 
    Set GetEstadoTPCta = rs
 
  Set rs.ActiveConnection = Nothing
  Set rs = Nothing

End Function


' Modificado por RIRO el 20130411
Public Function GetUltEstadoTP(ByVal nNumsol As Long, ByVal scperscod As String, ByVal nProd As Integer, ByVal nMon As Integer) As ADODB.Recordset
  Dim rs As ADODB.Recordset
   
  sSql = "Select  nNumSolicitud,cMovNro,cPersCod,nProducto,nMoneda,cCtaCod,nMonto,nPlazo,ntasa,nestado,cComentario,NEXTORNO,bpermanente=isnull(bpermanente,0), k.cConsDescripcion from captasaespecial c left join Constante k on k.nConsCod = LEFT(c.cSubProducto, 4) and k.nConsValor = right(c.cSubProducto, 1) where nnumsolicitud=" & nNumsol & " and nproducto=" & nProd & " and cperscod='" & scperscod & "' and nmoneda=" & nMon
  sSql = sSql & " and left(cmovnro,14)+cast(nnumsolicitud as varchar(12)) in (select max(left(c.cmovnro,14)+cast(c.nnumsolicitud as varchar(12))) from captasaespecial c where  nnumsolicitud=" & nNumsol & " and nproducto=" & nProd & " and nmoneda=" & nMon & ")"
   
'  sSql = "Select  nNumSolicitud,cMovNro,cPersCod,nProducto,nMoneda,cCtaCod,nMonto,nPlazo,ntasa,nestado,cComentario,NEXTORNO,bpermanente=isnull(bpermanente,0) from captasaespecial where nnumsolicitud=" & nNumsol & " and nproducto=" & nProd & " and cperscod='" & scperscod & "' and nmoneda=" & nMon
'  sSql = sSql & " and left(cmovnro,14)+cast(nnumsolicitud as varchar(12)) in (select max(left(c.cmovnro,14)+cast(c.nnumsolicitud as varchar(12))) from captasaespecial c where  nnumsolicitud=" & nNumsol & " and nproducto=" & nProd & " and nmoneda=" & nMon & ")"

  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseClient
  rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
  Set rs.ActiveConnection = Nothing
  Set GetUltEstadoTP = rs
  Set rs = Nothing
 
End Function


Public Function GetDatosTasaEspecialSol() As ADODB.Recordset
Dim rs As ADODB.Recordset

sSql = " SELECT P.cPersNombre, cSubProducto = K.cConsDescripcion, cTasaTarif = "
sSql = sSql & " dbo.ConvierteTNAaTEA(c.nTasaTarif), cTeaSoli = "
sSql = sSql & " dbo.ConvierteTNAaTEA(c.nTasaSolicitada) , cMoneda = CASE WHEN c.nMoneda = 1 THEN 'S/.' ELSE 'US$' END, "
sSql = sSql & " c.nMonto , c.nPlazo,c.cComentario, c.cPersCod, c.nMoneda, c.cMovnro, c.nProducto, c.nNumSolicitud, c.csubproducto, c.nTasaSolicitada"
sSql = sSql & " FROM CAPTASAESPECIAL c  JOIN Persona P ON p.cPersCod = c.cPersCod Left Join Constante K ON (k.nConsCod "
sSql = sSql & " = left(c.csubproducto,4) and  k.nConsValor = RIGHT(c.csubproducto, 1)) Where  left(c.cmovnro,14)+cast(c.nnumsolicitud "
sSql = sSql & " as varchar(12))  in  ( select max(left(c.cmovnro,14))+cast(c.nnumsolicitud as varchar(12)) from CAPTASAESPECIAL c    "
sSql = sSql & " group by c.nnumsolicitud  ) and (nestado=0 or (nestado=3 and nextorno=1) or (nestado=3 and nextorno=4)) order by c.nNumSolicitud "

'sSql = " SELECT P.cPersNombre, cProducto = K.cConsDescripcion, nTasa=case when c.nestado<>3  then  dbo.ConvierteTNAaTEA(c.nTasa) else c.ntasa  end  , cMoneda = CASE WHEN c.nMoneda = 1 THEN 'S/.' ELSE 'US$' END, "
'sSql = sSql & " c.nMonto , c.cComentario, c.cPersCod, c.nMoneda, c.nPlazo, c.cMovnro, c.nProducto, c.nNumSolicitud "
'sSql = sSql & " FROM  CAPTASAESPECIAL c "
'sSql = sSql & " JOIN Persona P ON p.cPersCod = c.cPersCod "
'sSql = sSql & " Left Join "
'sSql = sSql & " Constante K ON k.nConsValor = c.nProducto and k.nconscod=" & gProducto
'sSql = sSql & " Where "
'sSql = sSql & " left(c.cmovnro,14)+cast(c.nnumsolicitud as varchar(12))  in "
'sSql = sSql & " ( select max(left(c.cmovnro,14))+cast(c.nnumsolicitud as varchar(12)) from CAPTASAESPECIAL c "
'sSql = sSql & "   group by c.nnumsolicitud  ) "
'sSql = sSql & "   and (nestado=0 or (nestado=3 and nextorno=1) or (nestado=3 and nextorno=4)) "
'sSql = sSql & " order by c.nNumSolicitud "


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
rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetDatosTasaEspecialSol = rs
Set rs = Nothing


End Function


Public Function GetDatosCapTasaEspecial(ByVal nEstado As Integer) As ADODB.Recordset
Dim rs As ADODB.Recordset
sSql = "Select P.cPersNombre, cProducto = K.cConsDescripcion, T.nTasa, cMoneda = CASE WHEN T.nMoneda = 1 THEN 'S/.' ELSE 'US$' END, " _
    & "T.nMonto, T.cComentario, T.cPersCod, T.nMoneda, T.nPlazo, T.cMovNro, T.nProducto, T.nNumSolicitud From CapTasaEspecial T " _
    & "JOIN Persona P ON T.cPersCod = P.cPersCod LEFT JOIN Constante K ON T.nProducto = K.nConsValor Where " _
    & "T.cMovNro IN (Select Max(cMovNro) From CapTasaEspecial T1 Where T1.cPersCod = T.cPersCod And " _
    & "T1.nProducto = T1.nProducto And T1.nMoneda = T.nMoneda And T1.nNumSolicitud = T.nNumSolicitud) " _
    & "And T.nEstado = " & nEstado & " And K.nConsCod = " & gProducto
        
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
Set GetDatosCapTasaEspecial = rs
Set rs = Nothing
End Function

Public Function GetCargosCartasAI(ByVal nTipo As Integer, Filtro1 As String, Filtro2 As String, Filtro3 As String, psCodUser As String) As Boolean
Dim sSql As String, RSTMP As ADODB.Recordset
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
  
  sSql = "CREATE TABLE #TEMPCARTA" & psCodUser & "(COD int IDENTITY(1, 1) NOT NULL,DATO VARCHAR(2000),CARTA VARCHAR(10)) "
  Call Con.Ejecutar(sSql)
  
  sSql = "  INSERT into #tempcarta" & psCodUser & "(DATO ,CARTA ) "
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
   sSql = sSql & "SELECT DATO,CARTA FROM CARTASMREYNA ORDER BY NRO "

  Call Con.Ejecutar(sSql)
  
   sSql = " SELECT J.COD,J.DATO1,J.CARTA1,F.COD,DATO2=ISNULL(F.DATO2,''),ISNULL(F.CARTA2,0) "
   sSql = sSql & "     From "
   sSql = sSql & "     (SELECT  COD,DATO1=DATO,CARTA1=CARTA FROM #tempcarta" & psCodUser
   sSql = sSql & "      WHERE COD%2<>0 ) J "
   sSql = sSql & "     Left Join "
   sSql = sSql & "     (SELECT  COD,DATO2=DATO,CARTA2=CARTA FROM #tempcarta" & psCodUser
   sSql = sSql & "      WHERE COD%2=0  ) F ON F.COD=(J.COD+1) "
'   sSQL = sSQL & "     ORDER BY J.CARTA1,F.CARTA2 "
   RSTMP.Open sSql, Con.ConexionActiva, adOpenStatic, adLockOptimistic, adCmdText
    
    
    
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
    
   sSql = sSql & " drop table #tempcarta" & psCodUser
    
  
 Con.CierraConexion
 Set Con = Nothing

End Function
'JACA 20110610**********************************************************
Public Sub insertarOperacionInusual(ByVal psPersCod As String, ByVal pdFecha As Date, ByVal psAgencia As String, ByVal psMotivo As String, ByVal pnResultado As Integer, ByVal pnNroROS As Long, ByVal pnMovNro As String)

    sSql = "exec stp_ins_OpeInusual'" & psPersCod & "','" & Format(pdFecha, "yyyymmdd") & "','" & psAgencia & "','" & psMotivo & "'," & pnResultado & "," & pnNroROS & ",'" & pnMovNro & "'"
    dbCmact.Execute sSql

End Sub
Public Function getOperacionInusualPersona(ByVal psPersCod As String, ByVal psMovNro As String) As Recordset
    
    Dim rs As ADODB.Recordset
    sSql = "exec stp_sel_OpeInusual_Persona'" & psPersCod & "','" & psMovNro & "'"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set getOperacionInusualPersona = rs
    Set rs = Nothing

End Function
Public Sub actualizarOperacionInusualPersona(ByVal psMovNro As String)

    sSql = "exec stp_upd_OpeInusual_Persona'" & psMovNro & "'"
    dbCmact.Execute sSql

End Sub

'JACA END
'JACA 20110727**********************************************************
Public Sub insertarSectoresDJSujetosObligados(ByVal psCodCIIU As String, ByVal psMovNro As String)

    sSql = "exec stp_ins_SectoresDJSujetosObligados '" & psCodCIIU & "','" & psMovNro & "'"
    dbCmact.Execute sSql

End Sub
Public Function getSectoresDJSujetosObligados() As Recordset

    Dim rs As ADODB.Recordset
    sSql = "exec stp_sel_SectoresDJSujetosObligados"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set getSectoresDJSujetosObligados = rs
    Set rs = Nothing

End Function
Public Sub actualizarSectoresDJSujetosObligados()

    sSql = "exec stp_upd_SectoresDJSujetosObligados "
    dbCmact.Execute sSql

End Sub
Public Function getSectorDJSujetoObligado(ByVal psCodCIIU As String) As Boolean

    Dim rs As ADODB.Recordset
    sSql = "exec stp_sel_SectorDJSujetoObligado'" & psCodCIIU & "'"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    If Not (rs.EOF And rs.BOF) Then
        getSectorDJSujetoObligado = True
    Else
        getSectorDJSujetoObligado = False
    End If
    
    Set rs = Nothing

End Function
Public Function getDJSujetoObligado(ByVal psPersCod As String) As Boolean

    Dim rs As ADODB.Recordset
    sSql = "exec stp_sel_DJSujetoObligado'" & psPersCod & "'"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    If Not (rs.EOF And rs.BOF) Then
        getDJSujetoObligado = True
    Else
        getDJSujetoObligado = False
    End If
    
    Set rs = Nothing

End Function
'EJVG20120815 ***
'Public Sub insertarDJSujetoObligado(ByVal psPersCod As String, ByVal pdFecha As Date, ByVal psUser As String)
'
'    sSql = "exec stp_ins_DJSujetoObligado '" & psPersCod & "','" & Format(pdFecha, "yyyymmdd") & "','" & psUser & "'"
'    dbCmact.Execute sSql
'
'End Sub
Public Sub insertarDJSujetoObligado(ByVal psPersCod As String, ByVal pdFecha As Date, ByVal psUser As String, ByVal psAgeCod As String, ByVal pnSujetoObligado As Integer, ByVal pnOfCumplimiento As Integer)
    sSql = "Exec stp_ins_DJSujetoObligado '" & psPersCod & "','" & Format(pdFecha, "yyyymmdd hh:mm:ss") & "','" & psUser & "','" & psAgeCod & "'," & pnSujetoObligado & "," & pnOfCumplimiento
    dbCmact.Execute sSql
End Sub
'END EJVG *******
'JACA END

'*** PEAC 20130226
Public Sub MantenimientoGiroDestinatario(ByVal sCuenta As String, ByVal sDestinatario As String, ByVal sReferencia As String, ByVal pNumDoc As String, ByVal pnMovNro As Long)
    sSql = "exec stp_upd_GiroDestinatario '" & sCuenta & "','" & sDestinatario & "','" & sReferencia & "','" & pNumDoc & "'," & pnMovNro
    dbCmact.Execute sSql
End Sub


' Agregado por RIRO el 201303141113 **
Public Function ObtenerPermisoTea(Grupo As String, Agencia As String) As ADODB.Recordset

 Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorObtenerPermisoTea
    
        sSql = sSql & "stp_sel_PermisoTea '" & Grupo & "', '" & Agencia & "' "
        
        Set oConecta = New COMConecta.DCOMConecta
                
        oConecta.AbreConexion
        Set ObtenerPermisoTea = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
                
        Exit Function
    
ErrorObtenerPermisoTea:
    Set ObtenerPermisoTea = Nothing
    Err.Raise Err.Number, "ObtenerPermisoTea", "Error producido al obtener el nivel de aprobacion del usuario, contactarse con el area de TI"

End Function

Public Sub EditarGrupoAgencia(ByVal nTea As Double, ByVal nIdGrupo As Integer, nDias As Double, nSubProductos As Integer, _
ByVal User As String, ByVal dFecSis As String)

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
        
    On Error GoTo ErrorEditarGrupoAgencia
              
        sSql = "stp_upd_ActualizarGrupoAgencia " & nTea & ", '" & dFecSis & "', '" & User & "', "
        sSql = sSql & nDias & ", " & nSubProductos & ", " & nIdGrupo
            
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.CargaRecordSet (sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Sub
    
ErrorEditarGrupoAgencia:

    Err.Raise Err.Number, "EditarGrupoAgencia", "Error producido al editar una TEA, contactarse con el area de TI"

End Sub

Public Function GetAgenciaTea(ByVal index As Integer) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorGetAgenciaTea
    
        sSql = "stp_sel_AgenciaTea " & index
               
        Set oConecta = New COMConecta.DCOMConecta
                
        oConecta.AbreConexion
        Set GetAgenciaTea = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
                
        Exit Function
    
ErrorGetAgenciaTea:
    Set GetAgenciaTea = Nothing
    Err.Raise Err.Number, "GetAgenciaTea", "Error producido al obtener la TEA, contactarse con el area de TI"

End Function

Public Sub QuitarGrupoAgencia(ByVal nIdDominio As Integer, ByVal dFecSis As String)

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
        
    On Error GoTo ErrorQuitarGrupoAgencia
    
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        
        'dFecSis = Format(dFecSis, "yyyyMMdd")
        sSql = "update GruposDominio set nEstado = 0, dfechafin = '" & dFecSis & "' where nGrupDominio = " & nIdDominio
        oConecta.CargaRecordSet (sSql)
        sSql = "update GrupoAgencia set nEstado = 0, dfechafin = '" & dFecSis & "', dUltimoCambio = '" & dFecSis _
                                                                & "' where nGrupDominio = " & nIdDominio
        
        oConecta.CargaRecordSet (sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Sub
        
ErrorQuitarGrupoAgencia:
    Err.Raise Err.Number, "QuitarGrupoAgencia", "Error producido al inhabilitar el Nivel de Aprobacion, contactarse con el area de TI"

End Sub

Public Function InsertarGrupo(ByVal grupoNuevo As String, ByVal dFecSis As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
        
    On Error GoTo ErrorInsertarGrupo
    
        'dFecSis = Format(dFecSis, "yyyyMMdd")
        sSql = "stp_ins_GruposDominio '"
        sSql = sSql & grupoNuevo & "', '" & dFecSis & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set InsertarGrupo = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
    
ErrorInsertarGrupo:

    Set InsertarGrupo = Nothing
    Err.Raise Err.Number, "Buscar Convenio ", "Error producido al insertar un nuevo Nivel de aprobacion a la Base de Datos, contactarse con el area de TI"

End Function

Public Sub InsertarGrupoAgencia(ByVal idAgencia As String, ByVal idGrupo As Integer, ByVal TEAAdicional As Double, _
nDias As Double, nSubProductos As Integer, ByVal User As String, ByVal dFecSis As String)

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
        
    On Error GoTo ErrorInsertarGrupoAgencia
        
        sSql = "stp_ins_grupoAgencia '"
        sSql = sSql & idAgencia & "', "
        sSql = sSql & idGrupo & ", "
        sSql = sSql & TEAAdicional & ", "
        sSql = sSql & nDias & ", "
        sSql = sSql & nSubProductos & ", '"
        sSql = sSql & User & "', '"
        sSql = sSql & dFecSis & "'"

        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.CargaRecordSet (sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Sub
    
ErrorInsertarGrupoAgencia:

    Err.Raise Err.Number, "Buscar Convenio ", "Error producido al insertar un Nivel de aprobacion, contactarse con el area de TI"

End Sub

Public Function GetAgencia() As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorGetAgencia
    
        sSql = "select cAgeCod, cAgeDescripcion from Agencias where nestado = 1"
                
        Set oConecta = New COMConecta.DCOMConecta
                
        oConecta.AbreConexion
        Set GetAgencia = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
                
        Exit Function
    
ErrorGetAgencia:
    Set GetAgencia = Nothing
    Err.Raise Err.Number, "Buscar Convenio ", "Error producido al ontener las agencias de la BD, contactarse con el area de TI"

End Function

Public Function GetGruposDominio() As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorGetGruposDominio
    
        sSql = "select g.nGrupDominio, g.cNombreDominio, Ver='Ver',g.nEstado, g.dFechaInicio, g.dFechaFin  from GruposDominio g where g.nEstado=1"
                
        Set oConecta = New COMConecta.DCOMConecta
                
        oConecta.AbreConexion
        Set GetGruposDominio = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
                
        Exit Function
    
ErrorGetGruposDominio:
    Set GetGruposDominio = Nothing
    Err.Raise Err.Number, "Buscar Convenio ", "Error producido al obtener los Niveles de aprobacion activos, contactarse con el area de TI"

End Function

'************* fin riro
'RECO20140312 ERS174-2013************************************************
Public Function RecuperaDatosRelacGiro(ByVal psCtaCod As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaDatosRelacGiro
    
        sSql = "stp_sel_RecuperaDatosRelacGiro '" & psCtaCod & "'"
                
        Set oConecta = New COMConecta.DCOMConecta
                
        oConecta.AbreConexion
        Set RecuperaDatosRelacGiro = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
                
        Exit Function
    
ErrorRecuperaDatosRelacGiro:
    Set RecuperaDatosRelacGiro = Nothing
    Err.Raise Err.Number, "Recupera Datos", "Error producido al obtener los datos de las relaciones - Giro"

End Function
'RECO FIN ***************************************************************
'RECO20140312 ERS008-2014************************************************
Public Function RecuperaCodigoUltimoTarifarioGiro() As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaUltimoCodigoGiro
        sSql = "stp_sel_RecuperaultimoCodigoTarifarioGiros"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set RecuperaCodigoUltimoTarifarioGiro = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorRecuperaUltimoCodigoGiro:
    Set RecuperaCodigoUltimoTarifarioGiro = Nothing
    Err.Raise Err.Number, "Recupera Datos", "Error producido al obtener ultimo codigo - Giro"
End Function

Public Sub RegistraTarifarioGiros(ByVal psGiroTarCod As String, ByVal psFecha As String, ByVal pnMoneda As Integer, ByVal pnDesde As Double, ByVal pnHasta As Double, ByVal pnTipo As Integer, ByVal pnValor As Double, ByVal psUserReg As String, ByVal pnEstado As Integer)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorRecuperaUltimoCodigoTarifarioGiro
        sSql = "stp_ins_RegistraTarifarioGiros '" & psGiroTarCod & "','" & psFecha & "'," & pnMoneda & "," & pnDesde & "," & pnHasta & "," & pnTipo & "," & pnValor & ",'" & psUserReg & "'," & pnEstado
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.CargaRecordSet (sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Sub
ErrorRecuperaUltimoCodigoTarifarioGiro:
    Err.Raise Err.Number, "Recupera Datos", "Error producido al intentar registrar tarifario - Giro"
End Sub

Public Sub RegistraTarifarioGirosDet(ByVal psGiroTarCod As String, ByVal psAgeCod As String, ByVal pnEstado As Integer)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorRecuperaUltimoCodigoTarifarioGiroDet
        sSql = "stp_ins_RegistraTarifarioGirosDet '" & psGiroTarCod & "','" & psAgeCod & "'," & pnEstado
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.CargaRecordSet (sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Sub
ErrorRecuperaUltimoCodigoTarifarioGiroDet:
    Err.Raise Err.Number, "Recupera Datos", "Error producido al intentar registrar tarifario detalle - Giro"
End Sub
Public Function RecuperaDatosTarifarioGiros() As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaUltimoCodigoGiro
        sSql = "stp_sel_RecuperaDatosTarifarioGiros"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set RecuperaDatosTarifarioGiros = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorRecuperaUltimoCodigoGiro:
    Set RecuperaDatosTarifarioGiros = Nothing
    Err.Raise Err.Number, "Recupera Datos", "Error producido al recuperar datos - Giro"
End Function
Public Function ListaAgenciaTarGiros(ByVal psGiroTarCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaUltimoCodigoGiro
        sSql = "stp_sel_ListaAgenciaTarGiros '" & psGiroTarCod & "'"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ListaAgenciaTarGiros = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorRecuperaUltimoCodigoGiro:
    Set ListaAgenciaTarGiros = Nothing
    Err.Raise Err.Number, "Recupera Datos", "Error producido al recuperar datos - Giro"
End Function

Public Function RecuperaDatosGirosServAdic() As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaUltimoCodigoGiro
        sSql = "stp_sel_RecuperaDatosGirosServAdic"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set RecuperaDatosGirosServAdic = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorRecuperaUltimoCodigoGiro:
    Set RecuperaDatosGirosServAdic = Nothing
    Err.Raise Err.Number, "Recupera Datos", "Error producido al recuperar datos - Giro"
End Function
Public Sub ActualizaRecuperaDatosGirosServAdic(ByVal psGiroTarServAdicCod As String, ByVal pnMontoMN As Double, ByVal pnMontoME As Double)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorRecuperaUltimoCodigoTarifarioGiroDet
        sSql = "stp_upd_ActualizaRecuperaDatosGirosServAdic '" & psGiroTarServAdicCod & "'," & pnMontoMN & "," & pnMontoME
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.CargaRecordSet (sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Sub
ErrorRecuperaUltimoCodigoTarifarioGiroDet:
    Err.Raise Err.Number, "Recupera Datos", "Error producido al intentar registrar tarifario detalle - Giro"
End Sub
Public Function RecuperaValorComisionTarGiro(ByVal pnTpoOpe As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaUltimoCodigoGiro
        sSql = "stp_sel_RecuperaValorComisionTarGiro " & pnTpoOpe
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set RecuperaValorComisionTarGiro = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorRecuperaUltimoCodigoGiro:
    Set RecuperaValorComisionTarGiro = Nothing
    Err.Raise Err.Number, "Recupera Datos", "Error producido al recuperar datos - Giro"
End Function
Public Function HistorialMovimientoGiros(ByVal psPersCod As String, ByVal pntpoBusqueda As Integer, ByVal psFecIni As String, ByVal psFecFin As String, ByVal pnMoneda As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaUltimoCodigoGiro
        sSql = "stp_sel_HistorialMovimientoGiros '" & psFecIni & "','" & psFecFin & "','" & psPersCod & "'," & pnMoneda & "," & pntpoBusqueda
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set HistorialMovimientoGiros = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorRecuperaUltimoCodigoGiro:
    Set HistorialMovimientoGiros = Nothing
    Err.Raise Err.Number, "Recupera Datos", "Error producido al recuperar datos - Giro Movimientos"
End Function
Public Function ObtieneTotalesMovimientoGiro(ByVal psPersCod As String, ByVal pntpoBusqueda As Integer, ByVal psFecIni As String, ByVal psFecFin As String, ByVal pnMoneda As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtieneTotalesMovimientoGiro
        sSql = "stp_sel_ObtieneTotalesMovimientoGiro '" & psFecIni & "','" & psFecFin & "','" & psPersCod & "'," & pnMoneda & "," & pntpoBusqueda
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ObtieneTotalesMovimientoGiro = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorObtieneTotalesMovimientoGiro:
    Set ObtieneTotalesMovimientoGiro = Nothing
    Err.Raise Err.Number, "Recupera Datos", "Error producido al recuperar datos - Giro Movimientos"
End Function
Public Function RecuperaDatosTarifarioEditar(ByVal psGiroTarCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaDatosTarifarioEditar
        sSql = "stp_sel_RecuperaDatosTarifarioEditar '" & psGiroTarCod & "'"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set RecuperaDatosTarifarioEditar = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorRecuperaDatosTarifarioEditar:
    Set RecuperaDatosTarifarioEditar = Nothing
    Err.Raise Err.Number, "Recupera Datos", "Error producido al recuperar datos - Giro"
End Function
Public Function RecuperaDatosTarifarioDetEditar(ByVal psGiroTarCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaDatosTarifarioDetEditar
        sSql = "stp_sel_RecuperaDatosTarifarioDetEditar '" & psGiroTarCod & "'"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set RecuperaDatosTarifarioDetEditar = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorRecuperaDatosTarifarioDetEditar:
    Set RecuperaDatosTarifarioDetEditar = Nothing
    Err.Raise Err.Number, "Recupera Datos", "Error producido al recuperar datos - Giro"
End Function
Public Sub EliminaAgeTarifarioGiro(ByVal psGiroTarCod As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorRecuperaEliminaAgeTarifarioGiro
        sSql = "stp_del_EliminaAgeTarifarioGiro '" & psGiroTarCod & "'"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.CargaRecordSet (sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Sub
ErrorRecuperaEliminaAgeTarifarioGiro:
    Err.Raise Err.Number, "Elimina Datos", "Error producido al intentar eliminar agencias - Giro"
End Sub
Public Sub ActualizaTarifarioGiro(ByVal psGiroTarCod As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorActualizaTarifarioGiro
        sSql = "stp_upd_ActualizaTarifarioGiro '" & psGiroTarCod & "'"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.CargaRecordSet (sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Sub
ErrorActualizaTarifarioGiro:
    Err.Raise Err.Number, "Actualizar Datos", "Error producido al intentar Actualizar - Giro"
End Sub

Public Sub ActualizaParamComisionValorMin()
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorActualizaParamComisionValorMin
        sSql = "stp_upd_ActualizaParamComisionValorMin"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.CargaRecordSet (sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Sub
ErrorActualizaParamComisionValorMin:
    Err.Raise Err.Number, "Actualizar Datos", "Error producido al intentar Actualizar - Los Montos Mínimos de Comisión"
End Sub 'NAGL 20181006 Según RFC1807260001

Public Sub EliminaDestinatarioAnterior(ByVal psCtaCod As String, ByVal pnTpoClient As Integer)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorEliminaDestinatario
        sSql = "stp_del_EliminaDestinatarioAnterior '" & psCtaCod & "'," & pnTpoClient
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.CargaRecordSet (sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Sub
ErrorEliminaDestinatario:
    Err.Raise Err.Number, "Eliminar Destinatario", "Error producido al intentar eliminar destinatario - Giro"
End Sub
Public Function ActualizaDestinatarioAgeGiro(ByVal sCuenta As String, ByVal sAgenciaDest As String, ByVal psClaveGiro As String)
    Dim sClaveGiro As String
    Dim sSql As String
    Dim objEnc As New COMConecta.DCOMClasIni
    sClaveGiro = objEnc.Encripta(psClaveGiro)
    Set objEnc = Nothing
    sSql = "stp_upd_ActualizaDestinatarioAgeGiro '" & sCuenta & "','" & sAgenciaDest & "','" & sClaveGiro & "'"
    dbCmact.Execute sSql
End Function
Public Sub RegitraCambioDestinatarioGiro(ByVal psCtaCod As String, ByVal psAgeCodAct As String, ByVal psAgeCodAnt As String, ByVal pnTpoClient As Integer, ByVal psPersCodAct, ByVal psNumDOIAct As String, ByVal psPersCodAnt As String, ByVal psNumDOIAnt As String, ByVal psFechaReg As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorRegistroCambDest
        sSql = "stp_ins_RegitraCambioDestinatarioGiro '" & psCtaCod & "','" & psAgeCodAct & "','" & psAgeCodAnt & "'," & pnTpoClient & ",'" & psPersCodAct & "','" & psNumDOIAct & "','" & psPersCodAnt & "','" & psNumDOIAnt & "','" & Format(psFechaReg, "yyyyMMdd") & "'"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.CargaRecordSet (sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Sub
ErrorRegistroCambDest:
    Err.Raise Err.Number, "Cambio Destinatario", "Error producido al intentar realizar cambio destinatario - Giro"
End Sub
Public Function ObtieneValorTarifarioGiro(ByVal psAgeCod As String, ByVal pnMonto As Double, ByVal pnMoneda As Double) As ADODB.Recordset
       Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtieneValorTarifarioGiro
        sSql = "stp_sel_ObtieneValorTarifarioGiro '" & psAgeCod & "'," & pnMonto & "," & pnMoneda
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ObtieneValorTarifarioGiro = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorObtieneValorTarifarioGiro:
    Set ObtieneValorTarifarioGiro = Nothing
    Err.Raise Err.Number, "Recupera Datos", "Error producido al recuperar datos tarifario - Giro"
End Function
'RECO FIN********************************************************************
'FRHU 20140726 ERS106-2014
Public Function obtenerNivelesRiesgoVariables(ByVal pnTipo As Integer, Optional pnCodigo As Integer = 0) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorobtenerNivelesRiesgoVariables
        sSql = "stp_sel_NivelesRiesgoVariables " & pnTipo & "," & pnCodigo
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set obtenerNivelesRiesgoVariables = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorobtenerNivelesRiesgoVariables:
    Set obtenerNivelesRiesgoVariables = Nothing
    Err.Raise Err.Number, "Recupera Datos", "Error producido al recuperar las Variables"
End Function
Public Sub ActualizarValorNivelesDeRiesgos(ByVal pnTipo As Integer, Optional pnVarCod As Integer = 0, Optional pnVarSubCod As Integer = 0, Optional pnVarSubDetCod As Long = 0, Optional pnVarValor As Integer = 0)
    sSql = "stp_upd_NivelesRiesgoVariables " & pnTipo & "," & pnVarCod & "," & pnVarSubCod & "," & pnVarSubDetCod & "," & pnVarValor
    dbCmact.Execute sSql
End Sub
'FIN FRHU 20140726
'FRHU 20140915 ANEXO 01-ERS106-2014
Public Function obtenerPerfilDeRiesgoPorActividad(ByVal pnTipoPersona As Integer, ByVal pcOcupacion As String, ByVal pnMes As Integer, ByVal pcAnio As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorobtenerPerfilDeRiesgoPorActividad
        sSql = "exec stp_sel_NivelRiesgoPorActividad " & pnTipoPersona & ",'" & pcOcupacion & "'," & pnMes & ",'" & pcAnio & "'"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set obtenerPerfilDeRiesgoPorActividad = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorobtenerPerfilDeRiesgoPorActividad:
    Set obtenerPerfilDeRiesgoPorActividad = Nothing
    Err.Raise Err.Number, "Recupera Datos", "Error producido al recuperar el Perfil de Riesgo Por Actividad"
End Function
Public Function ValidarNivelDeRiesgo(ByVal pnTipoPersona As Integer, ByVal pcOcupacion As String, ByVal pnMes As Integer, ByVal pcAnio As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_ValidarNivelRiesgoPorActividad " & pnTipoPersona & ",'" & pcOcupacion & "'," & pnMes & ",'" & pcAnio & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ValidarNivelDeRiesgo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Sub EliminarNivelDeRiesgo(ByVal pnTipoPersona As Integer, ByVal pcOcupacion As String, ByVal pnMes As Integer, ByVal pcAnio As String)
    sSql = "exec stp_del_EliminarNivelRiesgoPorActividad " & pnTipoPersona & ",'" & pcOcupacion & "'," & pnMes & ",'" & pcAnio & "'"
    dbCmact.Execute sSql
End Sub
Public Sub RegistrarPerfilPorActividad(ByVal pnTipoPersona As Integer, ByVal pcOcupacion As String, ByVal pnMes As Integer, ByVal pcAnio As String, _
                                       ByRef pValPers() As Integer, ByRef pValPersPond() As Double, _
                                       ByRef pValProd() As Integer, ByRef pValProdPond() As Double, _
                                       ByRef pValCan() As Integer, ByRef pValCanPond() As Double, _
                                       ByRef pValJur() As Integer, ByRef pValJurPond() As Double, _
                                       ByRef pValComp() As Integer, ByRef pValCompPond() As Double, _
                                       ByRef pValTrans() As Integer, ByRef pValTransPond() As Double, _
                                       ByVal pnValPersona As Double, ByVal pnValProductos As Double, ByVal pnValCanales As Double, _
                                       ByVal pnValJurisdiccion As Double, ByVal pnValComportamiento As Double, ByVal pnValTransacciones As Double, _
                                       ByVal pnCalificacion As Double, ByVal pcNivelRiesgo As String)
    
    sSql = "exec stp_ins_PerfilRiesgo " & pnTipoPersona & ",'" & pcOcupacion & "'," & pnMes & ",'" & pcAnio & "',"
    sSql = sSql & pValPers(1) & "," & pValPers(2) & "," & pValPers(3) & "," & pValPers(4) & "," & pValPers(5) & "," & pValPers(6) & "," & pValPers(7) & "," & pValPers(8) & "," & pValPers(9) & "," & pValPers(10) & ","
    sSql = sSql & pValPersPond(1) & "," & pValPersPond(2) & "," & pValPersPond(3) & "," & pValPersPond(4) & "," & pValPersPond(5) & "," & pValPersPond(6) & "," & pValPersPond(7) & "," & pValPersPond(8) & "," & pValPersPond(9) & "," & pValPersPond(10) & ","
    sSql = sSql & pValProd(1) & "," & pValProd(2) & "," & pValProd(3) & "," & pValProd(4) & "," & pValProd(5) & "," & pValProd(6) & "," & pValProd(7) & "," & pValProd(8) & "," & pValProd(9) & ","
    sSql = sSql & pValProdPond(1) & "," & pValProdPond(2) & "," & pValProdPond(3) & "," & pValProdPond(4) & "," & pValProdPond(5) & "," & pValProdPond(6) & "," & pValProdPond(7) & "," & pValProdPond(8) & "," & pValProdPond(9) & ","
    sSql = sSql & pValCan(1) & "," & pValCan(2) & "," & pValCan(3) & "," & pValCan(4) & "," & pValCan(5) & "," & pValCan(6) & ","
    sSql = sSql & pValCanPond(1) & "," & pValCanPond(2) & "," & pValCanPond(3) & "," & pValCanPond(4) & "," & pValCanPond(5) & "," & pValCanPond(6) & ","
    sSql = sSql & pValJur(1) & "," & pValJur(2) & ","
    sSql = sSql & pValJurPond(1) & "," & pValJurPond(2) & ","
    sSql = sSql & pValComp(1) & "," & pValComp(2) & "," & pValComp(3) & "," & pValComp(4) & ","
    sSql = sSql & pValCompPond(1) & "," & pValCompPond(2) & "," & pValCompPond(3) & "," & pValCompPond(4) & ","
    sSql = sSql & pValTrans(1) & "," & pValTrans(2) & "," & pValTrans(3) & "," & pValTrans(4) & ","
    sSql = sSql & pValTransPond(1) & "," & pValTransPond(2) & "," & pValTransPond(3) & "," & pValTransPond(4) & ","
    sSql = sSql & pnValPersona & "," & pnValProductos & "," & pnValCanales & "," & pnValJurisdiccion & "," & pnValComportamiento & "," & pnValTransacciones & ","
    sSql = sSql & pnCalificacion & ",'" & pcNivelRiesgo & "'"
    
    dbCmact.Execute sSql
End Sub
'FIN FRHU 20140915
'GIRO HB----------
Public Function GetGiroBancaByInternet(ByVal sCuenta) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    sSql = "exec stp_sel_HB_giroBancaByInternet " & sCuenta
    oConecta.AbreConexion
    Set GetGiroBancaByInternet = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'-----------------
'CTI7 OpeV2**********************************************************************************
Public Sub ActualizarRelacionVoucherCaptacion(ByVal pnMovNroOpe As Long, _
                                              ByVal pnMovNroRVD As Long)
    Dim lsSQL As String
    lsSQL = "exec stp_upd_ActualizarCapVoucherDeposito_2  " & pnMovNroOpe & ", " & pnMovNroRVD & ""
    dbCmact.Execute lsSQL
    Exit Sub
End Sub
Public Sub ActualizaMovPendientesRend(ByVal nMovNroPend As Long, ByVal pnMonto As Currency)
    Dim lsSQL As String
        
    lsSQL = " UPDATE MovPendientesRend " _
          & " Set nSaldo = nSaldo - " & pnMonto & "" _
          & " Where nMovNro = " & nMovNroPend & ""
    
    dbCmact.CargaRecordSet lsSQL
    Exit Sub
End Sub
Public Function CapCargoCuentaAho(ByRef pMatAho As Variant, ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, ByVal pdFecSis As Date, _
            Optional nTipoDoc As TpoDoc = TpoDocNotaCargo, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional bOPExiste As Boolean = False, _
            Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
            Optional sCodCMAC As String = "", Optional sNomCmac As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "", Optional bCancelaciones As Boolean = False, _
            Optional pnMonedaCred As Moneda = 0, _
            Optional pbITFAplica As Boolean = True, Optional pnITFValor As Double = 0, Optional pbITFAsumido As Boolean = False, _
            Optional psItfOperacion As String = "", Optional pbRetiroEfectivo As Boolean = False, _
            Optional pnComisionValor As Double = 0) As Double
            'FRHU RQ14008 Se agrego: pbRetiroEfectivo

Dim rsCta As ADODB.Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nSaldoInac As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, i As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
Dim lnMontoOri As Double
Dim oGen As COMDConstSistema.DCOMGeneral
Dim lnTipCambioV As Currency
Dim lnTipCambioC As Currency
Dim lnTipCambio As Currency

'Dim clsMov As COMNContabilidad.NCOMContFunciones

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo

Set rsCta = GetDatosCuentaAho(sCuenta)

nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

Set oGen = New COMDConstSistema.DCOMGeneral

'GetTipCambio pdFecSis
lnTipCambioC = oGen.EmiteTipoCambio(pdFecSis, TCCompra)
lnTipCambioV = oGen.EmiteTipoCambio(pdFecSis, TCVenta)
lnTipCambio = oGen.EmiteTipoCambio(pdFecSis, TCFijoDia)

'RIRO 20200706 ****************
Dim sMensajeReactiva As String
sMensajeReactiva = oGen.LeeConstSistema("725")
If Not Len(Trim(sMensajeReactiva)) > 0 Then
    sMensajeReactiva = "Retiro Comisión Reactiva"
End If
'END RIRO *********************

If pnMonedaCred = 0 Then pnMonedaCred = Mid(sCuenta, 9, 1)

If pnMonedaCred <> Mid(sCuenta, 9, 1) Then
    lnMontoOri = nMonto
    If pnMonedaCred = gMonedaExtranjera And Mid(sCuenta, 9, 1) = gMonedaNacional Then
        nMonto = Round(nMonto * lnTipCambioV, 2)
    Else
        nMonto = Round(nMonto / lnTipCambioC, 2)
    End If
Else
    lnMontoOri = 0
End If

bTrans = True
Randomize
For i = 0 To Rnd(2000) * 1000
Next i

If Not ValidaSaldoCuenta(sCuenta, nMonto) Then 'Valida el saldo de la cuenta nuevamente
    Err.Raise 1000, "CapCargoCuentaAho", "Cuenta NO Posee Saldo Suficiente"
    CapCargoCuentaAho = 0
    Exit Function
End If
If bActivaCta Then
    If bInactiva Then
        nSaldoInac = nSaldoCnt
    End If
End If
'Calcula los intereses
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
If Not bCancelaciones Then
    pMatAho(3) = Format(nIntGanado, "#0.00")
Else
    pMatAho(4) = Format(nIntGanado, "#0.00")
End If

dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt - nMonto

ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp

nSaldoDisp = nSaldoDisp - nMonto
If Not bCancelaciones Then
    pMatAho(10) = nSaldoDisp
    pMatAho(11) = nSaldoCnt
Else
    pMatAho(12) = nSaldoDisp
    pMatAho(13) = nSaldoCnt
End If

'FRHU20140228 RQ14006
If pbRetiroEfectivo Then
    pMatAho(15) = nSaldoDisp
    pMatAho(16) = nSaldoCnt
    pMatAho(17) = Format(nIntGanado, "#0.00")
End If
'FIN FRHU20140228 RQ14006
ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True

If sNroDoc <> "" Then
    If nTipoDoc = TpoDocOrdenPago Then 'Si es con orden de pago
        If bOPExiste Then
            AgregaOrdenPagoEstado sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstCobrada
        Else
            AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMonto, gCapOPEstCobrada
        End If
        If nOperacion = gAhoRetOPCanje Then
            sMsgOpe = "Retiro OP Canje"
        ElseIf nOperacion = gAhoRetOP Then
            sMsgOpe = "Retiro OP"
        ElseIf nOperacion = gAhoOPCertificacion Then
            sMsgOpe = "Certificacion OP"
        
        'RIRO 20200630 ******
        ElseIf nOperacion = "100949" Then
            sMsgOpe = sMensajeReactiva
        'END RIRO ***********
            
        End If
    ElseIf nTipoDoc = TpoDocNotaCargo Then
        sMsgOpe = "Retiro NC"
    End If
Else
        'RIRO 20200630 ******
        If nOperacion = "100949" Then
            sMsgOpe = sMensajeReactiva
        'END RIRO ***********
        Else
            sMsgOpe = "Retiro Efectivo"
        End If
End If

'AgregaMov sMovNro, nOperacion, sGlosa

nMovNro = GetnMovNro(sMovNro)
AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
If bInactiva Then
    AgregaMovCap nMovNro, gAhoEstInacAct, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
    AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
    ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
    ActualizaInactivaAct (sCuenta) 'JUEZ 20150127
End If

'RIRO 20200611 Comisión Reactiva ****
If nOperacion = "100949" Then
    AgregaMovCapDet nMovNro, nOperacion, sCuenta, 1251, nMonto
Else
    If pnComisionValor > 0 Then
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto - pnComisionValor
         AgregaMovCapDet nMovNro, nOperacion, sCuenta, "3002", pnComisionValor
    Else
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
    End If
End If
'End RIRO 20200611 ******************

'AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto RIRO20200611 Comentado
'si el pago de credito es en diferente moneda a la cuenta de ahorros
Dim lnMontoCV As Double
If pnMonedaCred <> Mid(sCuenta, 9, 1) Then
    If Mid(sCuenta, 9, 1) = gMonedaExtranjera Then
        lnMontoCV = Round((lnMontoOri / lnTipCambioC) * lnTipCambio, 2)
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeCompra, lnMontoCV
        AgregaMovTipoCambio nMovNro, lnTipCambioC
        
        If lnTipCambioC < lnTipCambio Then
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - nMonto)
        Else
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - nMonto)
        End If
    Else
        lnMontoCV = Round(lnMontoOri * lnTipCambio, 2)
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeVenta, lnMontoCV
        AgregaMovTipoCambio nMovNro, lnTipCambioV
        
        If lnTipCambioV > lnTipCambio Then
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - Round(lnMontoOri * lnTipCambioV, 2))
        Else
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - Round(lnMontoOri * lnTipCambioV, 2))
        End If
    End If
End If
CapCargoCuentaAho = nSaldoCnt

If pbITFAplica And pnITFValor > 0 Then
    If pbITFAsumido Then
        AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
        AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFAsumido, pnITFValor
    Else
        ActualizaCargoCaptacion sCuenta, pnITFValor, pnITFValor, nIntGanado, dUltMov, sMovNro, bNumExtracto
        AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp - pnITFValor, nSaldoCnt - pnITFValor
        AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFCliente, pnITFValor
    End If
    
    nSaldoDisp = nSaldoDisp - pnITFValor
    If Not bCancelaciones Then
        pMatAho(10) = nSaldoDisp
        pMatAho(11) = nSaldoCnt - pnITFValor
    Else
        pMatAho(12) = nSaldoDisp
        pMatAho(13) = nSaldoCnt - pnITFValor
    End If
    'FRHU20140429 INCIDENTE: Se agrego para que en el vouche muestre el saldo correcto. cuando hace el retiro
    If pbRetiroEfectivo Then
        pMatAho(15) = nSaldoDisp
        pMatAho(16) = nSaldoCnt - pnITFValor
        pMatAho(17) = Format(nIntGanado, "#0.00")
    End If
    'FIN FRHU20140429
    CapCargoCuentaAho = nSaldoCnt - pnITFValor
End If
UltimaActualizacionCuenta sCuenta, sMovNro

bTrans = False
End Function
Public Sub UltimaActualizacionCuenta(ByVal sCuenta As String, ByVal sMovNro As String)
'Actualiza la fecha de ultima actualizacion de la cuenta de captaciones
Dim sSql As String
sSql = "Update Captaciones Set cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub
Public Function GetDatosCuentaAho(ByVal sCuenta As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset

sSql = "Exec sp_sel_DatosCuentaAhorro '" & sCuenta & "' , " & gCaptacEstado & " , " & gProductoCuentaTipo & " , " & gCaptacTipoTasa

Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosCuentaAho = rsCta
Set rsCta = Nothing
End Function
Public Function GetDatosCuentaPF(ByVal sCuenta As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, AG.cAgeDescripcion, " _
    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.nPlazo, A.nIntPag, A.dRenovacion, A.nApertura, " _
    & "A.nFormaRetiro, A.dAuxiliar, A.nDuplicado, T.cConsDescripcion cEstado, A.bBusCredPend, " _
    & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, T3.cConsDescripcion cRetiro, " _
    & "A.dUltCierreAnt, A.dAuxiliarAnt, ISNULL(CI.cCtaCodAbono,'') cCtaCodAbono,a.nformaretiro,c.calias,c.nfirmasmin FROM Producto P " _
    & "INNER JOIN Captaciones C INNER JOIN CaptacPlazoFijo A LEFT JOIN CaptacCtaAboIntPF CI " _
    & "ON A.cCtaCod = CI.cCtaCod ON C.cCtaCod = A.cCtaCod ON " _
    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Constante T3 ON A.nFormaRetiro = T3.nConsValor " _
    & "INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod WHERE P.cCtaCod = '" & sCuenta & "' AND " _
    & "T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    & "AND T2.nConsCod = " & gCaptacTipoTasa & " AND T3.nConsCod = " & gCaptacPFFormaRetiro


Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set GetDatosCuentaPF = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function
Public Function GetDatosCuentaCTS(ByVal sCuenta As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset

sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas,c.nfirmasmin ,C.nIntAcum, dUltCierre=( case when dbo.ultimodiames(dUltCierre)=1 then dUltCierre else dateadd(dd,1, dUltCierre) end )  , C.dApertura, AG.cAgeDescripcion, " _
    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.nSaldRetiro, A.nIntSaldo, PE.cPersNombre cInstitucion, " _
    & "T.cConsDescripcion cEstado, T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.cCodInst,C.cAlias " _
    & "FROM Producto P INNER JOIN Captaciones C INNER JOIN CaptacCTS A INNER JOIN " _
    & "Persona PE ON A.cCodInst = PE.cPersCod ON C.cCtaCod = A.cCtaCod ON " _
    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod " _
    & "WHERE P.cCtaCod = '" & sCuenta & "' AND T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    & "AND T2.nConsCod = " & gCaptacTipoTasa

Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosCuentaCTS = rsCta
Set rsCta = Nothing

End Function
Public Function ValidaSaldoCuenta(ByVal sCuenta As String, ByVal nMonto As Double) As Boolean
Dim nMoneda As Moneda
Dim rsCta As ADODB.Recordset
Dim nSaldo As Double, nMontoMinimo As Double

Set rsCta = GetDatosCuenta(sCuenta)
nSaldo = rsCta("nSaldoDisp")
rsCta.Close
Set rsCta = Nothing

nMoneda = CLng(Mid(sCuenta, 9, 1))
If nMoneda = gMonedaNacional Then
    nMontoMinimo = GetCapParametro(gSaldMinAhoMN)
Else
    nMontoMinimo = GetCapParametro(gSaldMinAhoME)
End If
If nSaldo - nMontoMinimo - nMonto >= 0 Then
    ValidaSaldoCuenta = True
Else
    ValidaSaldoCuenta = False
End If
End Function
Public Function GetDatosCuenta(ByVal sCuenta As String) As ADODB.Recordset
Dim rsCta As ADODB.Recordset
Dim nProd As Producto
nProd = CInt(Mid(sCuenta, 6, 3))

Select Case nProd
    Case gCapAhorros
        Set rsCta = GetDatosCuentaAho(sCuenta)
        
    Case gCapPlazoFijo
        Set rsCta = GetDatosCuentaPF(sCuenta)
        
    Case gCapCTS
        Set rsCta = GetDatosCuentaCTS(sCuenta)
        
End Select
Set GetDatosCuenta = rsCta
Set rsCta = Nothing

End Function
Public Function GetCapParametro(ByVal nParametro As CaptacParametro) As Double
Dim rsVar As ADODB.Recordset
Dim sSql As String
sSql = "SELECT nParValor FROM Parametro WHERE nParCod = " & nParametro & " " _
    & "And nParProd = " & gPrdParamCaptac
Set rsVar = New ADODB.Recordset
rsVar.CursorLocation = adUseClient
rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsVar.ActiveConnection = Nothing
If rsVar.EOF And rsVar.BOF Then
    GetCapParametro = 0
Else
    GetCapParametro = rsVar("nParValor")
End If
rsVar.Close
Set rsVar = Nothing
End Function

Public Function GetInteres(ByVal nCapital As Double, ByVal nTasa As Double, _
            ByVal nPlazo As Long, Optional nTipoInteres As TipoCalculoInteres = TpoCalcIntSimple) As Double
If nTipoInteres = TpoCalcIntSimple Then
    GetInteres = Round((nTasa / 36000) * nPlazo * nCapital, 2)
ElseIf nTipoInteres = TpoCalcIntCompuesto Then
    GetInteres = Round((((nTasa / 36000) + 1) ^ nPlazo - 1) * nCapital, 2)
End If
End Function

Public Sub AgregaOrdenPagoEstado(ByVal sCuenta As String, ByVal sNroDoc As String, _
        ByVal sMovNro As String, ByVal nMonto As Double, ByVal nEstadoOP As CaptacOrdPagoEstado)
Dim sSql As String
sSql = "UPDATE DocRecOP Set nMonto = " & nMonto & " WHERE nTpoDoc = " & TpoDocOrdenPago & " " _
    & "AND cNroDoc = '" & sNroDoc & "' And cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
sSql = "INSERT DocRecOPEst (nTpoDoc,cNroDoc,cCtaCod,cMovNro,nMonto,nEstado) " _
    & "VALUES (" & TpoDocOrdenPago & ",'" & sNroDoc & "','" & sCuenta & "','" & sMovNro & "'," & nMonto & "," & nEstadoOP & ")"
dbCmact.Execute sSql
End Sub

Public Sub AgregaCuentaDocumento(ByVal sCuenta As String, ByVal nTpoDoc As TpoDoc, _
        ByVal sNroDoc As String, ByVal sCodIF As String, ByVal sMovNro As String, _
        ByVal nMovNro As Long, Optional nMonto As Double = 0, _
        Optional nEstadoOP As CaptacOrdPagoEstado = gCapOPEstCobrada)
Dim sFecha As String
Dim sSql As String
sFecha = Mid(sMovNro, 1, 4) & "/" & Mid(sMovNro, 5, 2) & "/" & Mid(sMovNro, 7, 2) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
Select Case nTpoDoc
    Case TpoDocCheque
        sSql = "INSERT DocRecCapta (nTpoDoc,cNroDoc,cPersCod,cIFTpo,nMovNro,cCtaCod,nMonto) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "'," & nMovNro & ",'" & sCuenta & "'," & nMonto & ")"
    Case TpoDocOrdenPago
        sSql = "INSERT DocRecOP (nTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers,nEstado,cMovNro) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCuenta & "'," & nMonto & ",'" & sCodIF & "'," & nEstadoOP & ",'" & sMovNro & "')"
    Case TpoDocNotaAbono, TpoDocNotaCargo
End Select
dbCmact.Execute sSql
sSql = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
    & "VALUES (" & nMovNro & "," & nTpoDoc & ",'" & sNroDoc & "','" & sFecha & "')"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaEstadoCuenta(ByVal sCuenta As String, ByVal nEstado As CaptacEstado, _
        ByVal dFecha As Date, ByVal sMovNro As String)
'Actualiza el ultimo estado a la tabla de producto
Dim sSql As String
sSql = "Update Producto Set nPrdEstado = " & nEstado & ", " _
    & "dPrdEstado = '" & Format$(dFecha, "mm/dd/yyyy") & "' WHERE " _
    & "cCtaCod = '" & sCuenta & "'"
dbCmact.ConexionActiva.Execute sSql
'Agrega un registro mas de estado a la historia de estados de la cuenta
sSql = "INSERT CaptacEstado (cCtaCod,nPrdEstado,cMovNro) " _
    & "VALUES ('" & sCuenta & "'," & nEstado & ",'" & sMovNro & "')"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaInactivaAct(ByVal sCuenta As String)
Dim sSql As String
    sSql = "Update CaptacAhorros Set bInactiva = 0 WHERE cCtaCod = '" & sCuenta & "'"
    dbCmact.Execute sSql
End Sub

Public Sub AgregaMovTipoCambio(ByVal nMovNro As Long, ByVal pnTipoCambio As Currency)
Dim sSql As String

sSql = " INSERT MovTpoCambio (nMovNro, nMovTpoCambio) " _
     & " VALUES (" & nMovNro & "," & pnTipoCambio & ")"

dbCmact.Execute sSql
End Sub
'*******************************************************************************************


