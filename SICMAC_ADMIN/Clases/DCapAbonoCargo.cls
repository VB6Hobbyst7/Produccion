VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCapAbonoCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public dbCmact As Connection
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String
Dim sSql As String

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

Public Function AgregaNuevaCaptacion(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
        ByVal sAgencia As String, ByVal nTasa As Double, ByVal nSaldo As Double, _
        ByVal dFecha As Date, ByVal nFirmas As Integer, ByVal nPersoneria As PersPersoneria, _
        ByVal nTipoCuenta As ProductoCuentaTipo, ByVal sMovNro As String, _
        ByVal nTipoTasa As CaptacTipoTasa, Optional bCheque As Boolean = False, _
        Optional ByVal bOrdPag As Boolean = False, Optional nPlazo As Long = 0, _
        Optional nFormaRetiro As CaptacPFFormaRetiro = gCapPFFormRetMensual, _
        Optional bCtaAboInt As Boolean = False, Optional sCtaCodAbono As String = "", _
        Optional nPorcRetCTS As Double = 0, Optional sInstitucion As String = "") As String

Dim clsGen As DGeneral
Dim sCuenta As String, sFecha As String
Dim sFecUltCierre As String, sOrdPag As String, sCtaAbonoIntPF As String
Dim nSaldoDisp As Double, nSaldRetCTS As Double
Set clsGen = New DGeneral
sCuenta = clsGen.GeneraNuevaCuenta(sAgencia, nProducto, nMoneda)
Set clsGen = Nothing
sFecha = Format$(dFecha, gsFormatoFecha) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
sSql = "Insert Producto (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc) " _
    & "Values ('" & sCuenta & "'," & nTasa & "," & nSaldo & "," & gCapEstActiva & ",'" & sFecha & "',0)"
dbCmact.Execute sSql
nSaldoDisp = IIf(bCheque, 0, nSaldo)
sFecUltCierre = Format$(DateAdd("d", -1, dFecha), gsFormatoFecha)
sSql = "Insert Captaciones (cCtaCod,nSaldoDisp,nPersoneria,nFirmas,nIntAcum,dUltCierre,dApertura,nPrdCtaTpo,nPrdTasaInteres,cUltimaActualizacion) " _
    & "Values ('" & sCuenta & "'," & nSaldoDisp & "," & nPersoneria & "," & nFirmas & ",0,'" & sFecUltCierre & "','" & sFecha & "'," & nTipoCuenta & "," _
    & nTipoTasa & ",'" & sMovNro & "')"
dbCmact.Execute sSql
Select Case nProducto
    Case gCapAhorros
        sOrdPag = IIf(bOrdPag, "1", "0")
        sSql = "Insert CaptacAhorros (cCtaCod,nSaldoAnterior,bOrdPag,nSobregiro,dUltContacto) " _
            & "Values ('" & sCuenta & "',0," & sOrdPag & ",0,'" & sFecha & "')"
    Case gCapPlazoFijo
        sCtaAbonoIntPF = IIf(bCtaAboInt, "1", "0")
        sSql = "Insert CaptacPlazoFijo (cCtaCod,nPlazo,nIntPag,dRenovacion,nApertura,dAuxiliar,nFormaRetiro,nDuplicado,bAbonoIntCtaAho) " _
            & "Values ('" & sCuenta & "'," & nPlazo & ",0,'" & sFecha & "'," & nSaldo & ",'" & sFecha & "'," & nFormaRetiro & ",0," & sCtaAbonoIntPF & ")"
    Case gCapCTS
        nSaldRetCTS = Round(nSaldo * nPorcRetCTS / 100, 2)
        sSql = "Insert CaptacCTS (cCtaCod,nSaldRetiro,nIntSaldo,cCodInst) " _
            & "Values ('" & sCuenta & "'," & nSaldRetCTS & ",0,'" & sInstitucion & "')"
End Select
dbCmact.Execute sSql
If nProducto = gCapPlazoFijo And bCtaAboInt Then
    sSql = "Insert CaptacCtaAboIntPF (cCtaCod,cCtaCodAbono) " _
        & "Values ('" & sCuenta & "','" & sCtaCodAbono & "')"
    dbCmact.Execute sSql
End If
AgregaNuevaCaptacion = sCuenta
End Function

Public Sub AgregaNuevoProdPers(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona)
sSql = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " _
    & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ")"
dbCmact.Execute sSql
End Sub

Public Function GetDatosOrdenPago(ByVal sCuenta As String, ByVal sNroDoc As String) As ADODB.Recordset
Dim rsOP As ADODB.Recordset
sSql = "SELECT E.cMovNro, D.nTpoDoc, D.cNroDoc, D.cCtaCod, D.nMonto, P.cPersNombre, E.cEstado " _
    & "D.cIFCodPers FROM DocRecOPEst E INNER JOIN DocRecOP D INNER JOIN " & sDBPersona & "Persona P " _
    & "ON D.cIFCodPers = P.cPersCod ON E.nTpoDoc = D.nTpoDoc AND E.cNroDoc = D.cNroDoc AND E.cCtaCod " _
    & "= D.cCtaCod WHERE E.cMovNro IN (SELECT MAX(E1.cMovNro) FROM DocRecOPEst E WHERE E1.nTpoDoc = " _
    & "E.nTpoDoc AND E1.cNroDoc = E.cNroDoc AND E1.cCtaCod = E.cCtaCod)"
Set rsOP = New ADODB.Recordset
rsOP.CursorLocation = adUseClient
rsOP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsOP.ActiveConnection = Nothing
Set GetDatosOrdenPago = rsOP
Set rsOP = Nothing
End Function

Public Sub AgregaOrdenPagoEstado(ByVal sCuenta As String, ByVal sNroDoc As String, _
        ByVal sMovNro As String, ByVal nMonto As Double, ByVal nEstadoOP As CaptacOrdPagoEstado)
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
sFecha = Mid(sMovNro, 1, 4) & "/" & Mid(sMovNro, 5, 2) & "/" & Mid(sMovNro, 7, 2) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
Select Case nTpoDoc
    Case TpoDocCheque
        sSql = "INSERT DocRecCapta (nTpoDoc,cNroDoc,cPersCod,cIFTpo,cCtaCod,nMonto) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "','" & sCuenta & "'," & nMonto & ")"
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

Public Sub AgregaMov(ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
        ByVal sGlosa As String, Optional nMovEstado As MovEstado = gMovEstContabMovContable, _
        Optional nMovFlag As MovFlag = gMovFlagVigente)
sSql = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & nOperacion & "','" & sGlosa & "'," & nMovEstado & "," & nMovFlag & ")"
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
    & ", dUltCierre = '" & Format$(dUltMov, gsFormatoFecha) & "', cUltimaActualizacion = '" & sMovNro & "' " _
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
    & ", dUltCierre = '" & Format$(dUltMov, gsFormatoFecha) & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub UltimaActualizacionCuenta(ByVal sCuenta As String, ByVal sMovNro As String)
'Actualiza la fecha de ultima actualizacion de la cuenta de captaciones
sSql = "Update Captaciones Set cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaEstadoCuenta(ByVal sCuenta As String, ByVal nEstado As CaptacEstado, _
        ByVal dFecha As Date, ByVal sMovNro As String)
'Actualiza el ultimo estado a la tabla de producto
sSql = "Update Producto Set nPrdEstado = " & nEstado & ", " _
    & "dPrdEstado = '" & Format$(dFecha, gsFormatoFecha) & "' WHERE " _
    & "cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
'Agrega un registro mas de estado a la historia de estados de la cuenta
sSql = "INSERT CaptacEstado (cCtaCod,nPrdEstado,cMovNro) " _
    & "VALUES ('" & sCuenta & "'," & nEstado & ",'" & sMovNro & "')"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaEstadoDocRecEst(ByVal nTipoDoc As TpoDoc, ByVal sNroDoc As String, _
        ByVal sCodIF As String, ByVal sMovNro As String, ByVal nEstado As ChequeEstado)
'Actualiza el ultimo estado a la tabla de Documento Recibidos
sSql = "Update DocRec Set nEstado = " & nEstado & " WHERE " _
    & "nTpoDoc = '" & nTipoDoc & "' AND cNroDoc = '" & sNroDoc & "' AND cPersCod = '" & sCodIF & "'"
dbCmact.Execute sSql
'Agrega un registro mas de estado a la historia de estados del documento recibido
sSql = "INSERT DocRecEst (nTpoDoc,cNroDoc,cPersCod,cMovNro,nEstado) " _
    & "VALUES ('" & nTipoDoc & "','" & sNroDoc & "','" & sCodIF & "','" & sMovNro & "'," & nEstado & ")"
dbCmact.Execute sSql
End Sub

Public Function GetMovExtorno(ByVal sDatoBus As String, Optional nTipoBus As Integer = 0) As ADODB.Recordset
Dim rsMov As ADODB.Recordset
Dim sWhere As String
If nTipoBus = 0 Then
    sWhere = "M.cMovNro LIKE '%" & sDatoBus & "%'"
ElseIf nTipoBus = 1 Then
    sWhere = "M.cMovNro LIKE '" & Format$(gdFecSis, "yyyymmdd") & "%' And C.cCtaCod = '" & sDatoBus & "'"
End If
sSql = "Select M.cMovNro, O.cOpeDesc, C.cCtaCod, ISNULL(MD.nDocTpo,'') nDocTpo, ISNULL(MD.cDocNro,'') cDocNro, " _
    & "M.cMovDesc, C.cOpeCod, C.nMonto, M.nMovNro FROM MovDoc MD RIGHT JOIN Mov M INNER JOIN MovCap C " _
    & "INNER JOIN " & sDBComunes & "OpeTpo O ON C.cOpeCod = O.cOpeCod ON M.nMovNro = C.nMovNro " _
    & "ON MD.nMovNro = M.nMovNro WHERE " & sWhere & " AND C.cOpeCod NOT IN ('" & gAhoEstInacAct & "') " _
    & "AND M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagEliminado & "," & gMovFlagDeExtorno & ")" _
    & "AND M.nMovEstado IN (" & gMovEstContabMovContable & ") " _
    & "ORDER BY M.nMovNro"
Set rsMov = New ADODB.Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetMovExtorno = rsMov
End Function

Public Sub ActualizaEstadoMov(ByVal nMovNro As Long, ByVal nMovFlag As MovFlag)
sSql = "Update Mov Set nMovFlag = " & nMovFlag & " Where nMovNro = " & nMovNro
dbCmact.Execute sSql
End Sub

Public Function EsOrdenPagoEmitida(ByVal sCuenta As String, ByVal nNumOP As Long) As Boolean
Dim rsOP As ADODB.Recordset
Set rsOP = New ADODB.Recordset
rsOP.CursorLocation = adUseClient
sSql = "Select OP.cCtaCod, OP.nInicio, OP.nFin FROM Mov M INNER JOIN MovDocEmitidoRango OP ON " _
    & "M.nMovNro = OP.nMovNro WHERE M.nMovFlag <> " & gMovFlagExtornado & " AND OP.cCtaCod = '" _
    & sCuenta & "' And " & nNumOP & " Between OP.nInicio And OP.nFin"
rsOP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsOP.ActiveConnection = Nothing
If rsOP.EOF And rsOP.BOF Then
    EsOrdenPagoEmitida = False
Else
    EsOrdenPagoEmitida = True
End If
rsOP.Close
Set rsOP = Nothing
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

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
'************************************************************************
'************************************************************************
'************************************************************************
Public Function InsertaMovCont(ByVal pnMovNro As Long, ByVal pnMovMonto As Currency, ByVal pnMovCorrela As Integer, ByVal psMovParalelo As String) As Integer
    On Error GoTo InsertaMovContErr
    Dim sql As String
    InsertaMovCont = 1
    
    sql = "INSERT INTO MovCont (nMovNro, nMovMonto, nMovCorrela, cMovParalelo ) " _
         & "VALUES (" & pnMovNro & "," & pnMovMonto & "," & pnMovCorrela & ",'" & psMovParalelo & "')"
     
    dbCmact.Execute sql
    InsertaMovCont = 0
    Exit Function
InsertaMovContErr:
    Call RaiseError(MyUnhandledError, "DCapMov:InsertaMovContErr Method")
End Function
'##ModelId=3A848F08038A
Public Function InsertaMovCta(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psCtaContCod As String, ByVal pnMovImporte As Currency) As Integer
    On Error GoTo InsertaMovCtaErr
    Dim sql As String
    InsertaMovCta = 1
    
    sql = " INSERT INTO MOVCTA(nMovNro, nMovItem, cCtaContCod, nMovImporte )" _
        & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psCtaContCod & "'," & pnMovImporte & ") "
      
    dbCmact.Execute sql
    InsertaMovCta = 0
    Exit Function
InsertaMovCtaErr:
    Call RaiseError(MyUnhandledError, "DCapMov:InsertaMovCta Method")
End Function
Public Function InsertaMovObj(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As Long, ByVal psObjetoCod As TpoObjetos) As Integer
    On Error GoTo InsertaMovObjErr
    Dim sql As String
    InsertaMovObj = 1
     
    sql = "INSERT INTO MOVOBJ (nMovNro, nMovItem, nMovObjOrden, cObjetoCod) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & Format(psObjetoCod, "00") & "')"

    dbCmact.Execute sql
    InsertaMovObj = 0
    Exit Function
InsertaMovObjErr:
    Call RaiseError(MyUnhandledError, "DCapMov:InsertaMovObj Method")
End Function
Public Function InsertaMovObjAgenciaArea(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As Long, ByVal psAgeCod As String, ByVal psAreaCod As String) As Integer
    On Error GoTo InsertaMovObjAgenciaAreaErr
    Dim sql As String
    InsertaMovObjAgenciaArea = 1
    
    sql = "INSERT INTO MovObjAreaAgencia(nMovNro, nMovItem, nMovObjOrden, cAreaCod, cAgeCod) " _
         & "VALUES (" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & psAreaCod & "','" & psAgeCod & "')"
     
    dbCmact.Execute sql
    InsertaMovObjAgenciaArea = 0
    Exit Function
InsertaMovObjAgenciaAreaErr:
    Call RaiseError(MyUnhandledError, "DCapMov:InsertaMovObjAgenciaArea Method")
End Function
Public Function InsertaMovObjIF(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As String, ByVal psPersCod As String, ByVal psTipoIF As String, ByVal psCtaIfCod As String) As Integer
    On Error GoTo InsertaMovObjIFErr
    Dim sql As String
    InsertaMovObjIF = 1
     
    sql = "INSERT INTO MOVOBJIF(nMovNro, nMovItem, nMovObjOrden, cPersCod, cIFTpo, cCtaIfCod ) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & psPersCod & "','" & psTipoIF & "','" & psCtaIfCod & "')"
    
    dbCmact.Execute sql
    InsertaMovObjIF = 0
    Exit Function
InsertaMovObjIFErr:
    Call RaiseError(MyUnhandledError, "DCapMov:InsertaMovObjIF Method")
End Function
Public Function InsertaMovRef(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long) As Integer
    On Error GoTo InsertaMovRefErr
    Dim sql As String
    InsertaMovRef = 1
    
    sql = "INSERT INTO MovRef(nMovNro, nMovNroRef) " _
         & "VALUES (" & pnMovNro & "," & pnMovNroRef & ")"
     
    dbCmact.Execute sql
    InsertaMovRef = 0
    Exit Function
InsertaMovRefErr:
    Call RaiseError(MyUnhandledError, "DCapMov:InsertaMovContErr Method")
End Function
Public Function GeneraMovNro(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "") As String
    On Error GoTo GeneraMovNroErr
    Dim rs As ADODB.Recordset
    Dim sql As String
    Set rs = New ADODB.Recordset
    
    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & GetHoraServer, gsFormatoFechaHora) & "','" & psCodAge & "','" & psUser & "'"
    Else
       sql = "sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    Set rs = dbCmact.Execute(sql)
    If Not rs.EOF Then
        GeneraMovNro = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
GeneraMovNroErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function
Public Function GetcMovNro(ByVal pnMovNro As Long) As String
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

sql = "Select cMovNro From Mov where nMovNro ='" & pnMovNro & "'"
rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.EOF And Not rs.BOF Then
    GetcMovNro = rs!cMovNro
End If
rs.Close
Set rs = Nothing
End Function
Public Function ActualizaMovArendir(ByVal pnMovNro As Long, Optional ByVal psTpoArendir As ArendirTipo = gArendirTipoCajaGeneral, Optional ByVal psAreaCod As String, Optional ByVal psAgeCod As String, Optional ByVal psPersCod As String, Optional ByVal pnMovSaldo As Currency = -9999) As Integer
    On Error GoTo ActualizaMovARendirErr
    Dim sql As String
    Dim lsActualiza As String
    lsActualiza = ""
    ActualizaMovArendir = 1
    If psAgeCod <> "" Then
        lsActualiza = lsActualiza & " cAgeCod='" & psAgeCod & "',"
    End If
    If psAreaCod <> "" Then
        lsActualiza = lsActualiza & " cAreaCod='" & psAreaCod & "',"
    End If
    If psPersCod <> "" Then
        lsActualiza = lsActualiza & " cPersCod='" & psPersCod & "',"
    End If
    If pnMovSaldo <> -9999 Then
        lsActualiza = lsActualiza & " nMovSaldo=" & pnMovSaldo & ","
    End If
    'nMovSaldo
    lsActualiza = lsActualiza & " cTpoArendir='" & psTpoArendir & "',"
    lsActualiza = Mid(lsActualiza, 1, Len(lsActualiza) - 1)
    
    If lsActualiza = "" Then Exit Function
    sql = " Update MovARendir set " & lsActualiza & " Where nMovNro =" & pnMovNro & ""
    
    dbCmact.Execute sql
    ActualizaMovArendir = 0
    Exit Function
ActualizaMovARendirErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMovARendir Method")
End Function
Public Function ActualizaMovimiento(ByVal psMovNroNew As String, ByVal pnMovNroAnt As Long, Optional psMovEstado As MovEstado = gMovEstContabMovContable, _
                                    Optional plActualizaAsiento As Boolean = True, Optional pbInsertaMov As Boolean = True, Optional pbDeleteMovAnt As Boolean = True, _
                                    Optional ByVal psOpeCod As String = "", Optional ByVal psMovDesc As String = "", Optional ByVal lbUpdaMovRef As Boolean = True) As Integer
On Error GoTo ActualizaMovimientoErr
   Dim sql As String
   Dim lnNovNronew As Long
   Dim lnMovNronew As Long
   ActualizaMovimiento = 1
   sql = ""
   
   If plActualizaAsiento Then
      If pbInsertaMov Then
      
        sql = " INSERT INTO MOV (cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag )" _
              & " SELECT '" & psMovNroNew & "'," & IIf(psOpeCod = "", "cOpeCod", "'" & psOpeCod & "'") & "," _
            & IIf(psMovDesc = "", "cMovDesc", "'" & psMovDesc & "'") & ",'" & psMovEstado & "', nMovFlag  " _
            & " FROM Mov WHERE nMovNro =" & pnMovNroAnt & " "
        
        dbCmact.Execute sql
        lnMovNronew = GetnMovNro(psMovNroNew)
      Else
        lnMovNronew = GetnMovNro(psMovNroNew)
      End If
      sql = ""
      sql = " UPDATE MOVCONT        SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
      dbCmact.Execute sql
   End If
   sql = " UPDATE MOVGASTO       SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   dbCmact.Execute sql
   sql = "UPDATE MOVARQUEO      SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
   dbCmact.Execute sql
   sql = "UPDATE MOVDOC         SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
   dbCmact.Execute sql
   If lbUpdaMovRef Then
        sql = "UPDATE MOVREF         SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        sql = "UPDATE MOVREF         SET nMovNroRef =" & lnMovNronew & " Where nMovNroRef =" & pnMovNroAnt & " "
        dbCmact.Execute sql
   End If
   sql = "UPDATE MOVARENDIR     SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   dbCmact.Execute sql
   sql = "UPDATE MOVARENDIRSUST SET nMovNroRend=" & lnMovNronew & " where nMovNroSust =" & pnMovNroAnt & " "
   dbCmact.Execute sql
   sql = "UPDATE MOVOTROSITEM   SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   dbCmact.Execute sql
   sql = "UPDATE MOVCAJACHICA   SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   dbCmact.Execute sql
   sql = "UPDATE MOVCARTAFIANZA SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   dbCmact.Execute sql
   sql = "UPDATE MOVTPOCAMBIO   SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   dbCmact.Execute sql
   sql = "UPDATE MOVVIATICOS    SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   dbCmact.Execute sql
    
   If plActualizaAsiento Then
        sql = " INSERT MOVCTA (nMovNro, nMovItem, cCtaContCod, nMovImporte ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, cCtaContCod, nMovImporte FROM MOVCTA Where nMovNro =" & pnMovNroAnt & " "
        
        dbCmact.Execute sql
        
        sql = "INSERT MOVOBJ (nMovNro, nMovItem, nMovObjOrden, cObjetoCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cObjetoCod FROM MOVOBJ Where nMovNro =" & pnMovNroAnt & " "
        
        dbCmact.Execute sql
        
        sql = "UPDATE MOVOBJPers        SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        sql = "UPDATE MOVOBJIF          SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        sql = "UPDATE MOVOBJEfectivo    SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        sql = "UPDATE MOVOBJBS          SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        sql = "UPDATE MOVOBJAreaAgencia SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        sql = "UPDATE MOVME             SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        
        'sql = "UPDATE MOVOBJ            SET nMovNro =" & lnMovNroNew & " Where nMovNro =" & pnMovNroAnt & " "
        sql = "DELETE MOVOBJ            Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        'sql = "UPDATE MOVCTA            SET nMovNro =" & lnMovNroNew & " Where nMovNro =" & pnMovNroAnt & " "
        sql = "DELETE MOVCTA            Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        If pbDeleteMovAnt Then
            sql = "Update MOV set nMovFlag=" & gMovFlagEliminado & " where nMovNro =" & pnMovNroAnt & " "
            dbCmact.Execute sql
        End If
   End If
   ActualizaMovimiento = 0
   Exit Function
ActualizaMovimientoErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMovimiento Method")
End Function
Public Function GetHoraServer() As String
Dim sql As String
Dim rsH As New ADODB.Recordset
sql = "SELECT convert(varchar(10),getdate(),108) as sHora"
rsH.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
If Not rsH.EOF Then
   GetHoraServer = rsH!sHora
Else
   GetHoraServer = Format(Time, "hh:mm:ss")
End If
rsH.Close: Set rsH = Nothing
End Function
Public Function InsertaMovDoc(ByVal pnMovNro As Long, ByVal pnDocTpo As TpoDoc, ByVal psDocNro As String, ByVal psDocFecha As String) As Integer
    On Error GoTo InsertaMovDocErr
    Dim sql As String
    InsertaMovDoc = 1
    
    sql = "INSERT INTO MOVDOC(nMovNro, nDocTpo, cDocNro, dDocFecha ) " _
        & "VALUES (" & pnMovNro & "," & pnDocTpo & ",'" & psDocNro & "','" & Format(psDocFecha, gsFormatoFecha) & "')"
        
    dbCmact.Execute sql
    InsertaMovDoc = 0
    Exit Function
InsertaMovDocErr:
    Call RaiseError(MyUnhandledError, "DCapMov:InsertaMovDoc Method")
End Function
Public Function InsertaNotaAbonoCargoEst(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pnEstado As NotaCargoAbonoEstado, _
                                        ByVal psMovNro As String) As Integer
    On Error GoTo InsertaNotaAbonoCargoEstErr
    Dim sql As String
    InsertaNotaAbonoCargoEst = 1
    
    sql = " INSERT INTO NotaAbonoCargoEst(nDocTpo, cDocNro, nEstado, cMovNro)" _
        & " VALUES(" & pnTpoDoc & ",'" & psNroDoc & "'," & pnEstado & ",'" & psMovNro & "') "
    
    dbCmact.Execute sql
    InsertaNotaAbonoCargoEst = 0
    Exit Function
InsertaNotaAbonoCargoEstErr:
    Call RaiseError(MyUnhandledError, "DCapMov:InsertaNotaAbonoCargoEst Method")
End Function
Public Function ActualizaNotaAbonoCargo(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, Optional ByVal pnEstado As NotaCargoAbonoEstado = -1, _
                                    Optional ByVal pnMotivo As Integer = -1, Optional ByVal pnMonto As Currency = -999999, _
                                    Optional psObjetoCod As String = "", Optional psObjeto As String = "") As Integer
    On Error GoTo ActualizaNotaAbonoCargoErr
    Dim sql As String
    Dim lsActualiza As String
    ActualizaNotaAbonoCargo = 1
    lsActualiza = ""
    
    If pnMotivo <> -1 Then
        lsActualiza = lsActualiza + " nMotivoCod=" & pnMotivo & ", "
    End If
    If pnEstado <> -1 Then
        lsActualiza = lsActualiza + " nEstado=" & pnEstado & ", "
    End If
    If pnMonto <> -999999 Then
        lsActualiza = lsActualiza + " nMonto=" & pnMonto & ","
    End If
    If psObjetoCod <> "" Then
        lsActualiza = lsActualiza + " cObjetoCodPadre= '" & psObjetoCod & "',"
    End If
    If psObjeto <> "" Then
        lsActualiza = lsActualiza + " cObjetoCod='" & psObjeto & "',"
    End If
    If lsActualiza <> "" Then
        lsActualiza = Mid(lsActualiza, 1, Len(lsActualiza) - 1)
        sql = " UPDATE NotaAbonoCargo SET " & lsActualiza & " " _
            & " WHERE nDocTpo=" & pnTpoDoc & " AND cDocNro='" & psNroDoc & "'"
    End If
        
    dbCmact.Execute sql
    ActualizaNotaAbonoCargo = 0
    Exit Function
ActualizaNotaAbonoCargoErr:
    Call RaiseError(MyUnhandledError, "DCapMov:ActualizaNotaAbonoCargo Method")
End Function
Public Function ExtornaMovimiento(ByVal psMovNroNew As String, ByVal pnMovNroAnt As Long, _
                                Optional ByVal psOpeCod As String = "", Optional ByVal psMovDesc As String = "", _
                                Optional ByVal pbEliminaMov As Boolean, Optional pbInsertaNewMov As Boolean = True) As Integer
On Error GoTo ExtornaMovimientoErr
    Dim sql As String
    Dim lnMovNronew As Long
    Dim lsMovNroAnt As String
    Dim ldFechaAnt As Date
    Dim ldFechaExt As Date
    
    
    lsMovNroAnt = GetcMovNro(pnMovNroAnt)
    ldFechaAnt = CDate(Mid(lsMovNroAnt, 7, 2) & "/" & Mid(lsMovNroAnt, 5, 2) & "/" & Mid(lsMovNroAnt, 1, 4))
    ldFechaExt = CDate(Mid(psMovNroNew, 7, 2) & "/" & Mid(psMovNroNew, 5, 2) & "/" & Mid(psMovNroNew, 1, 4))
    
    
    ExtornaMovimiento = 1
    If pbInsertaNewMov Then
        sql = " INSERT INTO MOV (cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag )" _
            & " SELECT '" & psMovNroNew & "'," & IIf(psOpeCod = "", "cOpeCod", "'" & psOpeCod & "'") & ", " _
            & IIf(psMovDesc = "", "cMovDesc", "'" & Replace(psMovDesc, "''", "'") & "'") & ",nMovEstado, '" & gMovFlagDeExtorno & "'  " _
            & " FROM Mov WHERE nMovNro ='" & pnMovNroAnt & "' "
   
        dbCmact.Execute sql
    End If
    lnMovNronew = GetnMovNro(psMovNroNew)
    InsertaMovRef lnMovNronew, pnMovNroAnt
    If ldFechaAnt = ldFechaExt Or pbEliminaMov = True Then
        'ActualizaMov pnMovNroAnt, , , gMovFlagEliminado
    ElseIf ldFechaAnt <> ldFechaExt Or pbEliminaMov = False Then
        sql = ""
      
        sql = "INSERT MOVCONT (nMovNro, nMovMonto, nMovCorrela, cMovParalelo) " _
            & " SELECT nMovNro = " & lnMovNronew & ", nMovMonto, nMovCorrela, cMovParalelo FROM MovCont Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
       
        sql = "INSERT MOVDOC (nMovNro, nDocTpo, cDocNro, dDocFecha) " _
            & " SELECT nMovNro = " & lnMovNronew & ", nDocTpo, cDocNro, dDocFecha FROM MovDoc Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        
        sql = "INSERT MOVGASTO(nMovNro,cPersCod,cDestino) " _
            & " SELECT nMovNro = " & lnMovNronew & ",cPersCod,cDestino FROM MovGasto Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        
        sql = "INSERT MOVCTA (nMovNro, nMovItem, cCtaContCod, nMovImporte ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, cCtaContCod, nMovImporte * -1  FROM MOVCTA Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        
        sql = "INSERT MOVME (nMovNro, nMovItem, nMovMEImporte ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovMEImporte*-1 FROM MOVME Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        
        sql = "INSERT MOVOBJ (nMovNro, nMovItem, nMovObjOrden, cObjetoCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cObjetoCod FROM MOVOBJ Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        
        sql = "INSERT MOVOBJPers (nMovNro, nMovItem, nMovObjOrden, cPersCod) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cPersCod FROM MovObjPers Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        
        sql = "INSERT MOVOBJIF (nMovNro, nMovItem, nMovObjOrden, cPersCod, cCtaIfCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cPersCod, cCtaIfCod FROM MOVOBJIF Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        
        sql = "INSERT MOVOBJEfectivo (nMovNro, nMovItem, nMovObjOrden, cEfectivoCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cEfectivoCod FROM MOVOBJEfectivo Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        
        sql = "INSERT MOVOBJBS (nMovNro, nMovItem, nMovObjOrden, cBieSerCod, nMovCantidad ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cBieSerCod, nMovCantidad FROM MOVOBJBS Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        
        sql = "INSERT MOVOBJAreaAgencia (nMovNro, nMovItem, nMovObjOrden, cAreaCod, cAgeCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cAreaCod, cAgeCod FROM MOVOBJAreaAgencia Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        
        sql = "INSERT INTO MOVHABILITACION(nMovNro, cTranspCod, nMontoTrans) " _
            & " SELECT nMovNro=" & lnMovNronew & ", cTranspCod, nMontoTrans FROM MOVHABILITACION Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        
        sql = "INSERT INTO MOVHABDEVCAJERO (nMovNro, cUser, nMovImporte) " _
            & " SELECT nMovNro=" & lnMovNronew & ",cUser, nMovImporte FROM MOVHABDEVCAJERO Where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
        
        sql = "UPDATE MOV SET nMovFlag = '" & gMovFlagExtornado & "' where nMovNro =" & pnMovNroAnt & " "
        dbCmact.Execute sql
    End If
    
    ExtornaMovimiento = 0
    Exit Function
ExtornaMovimientoErr:
    Call RaiseError(MyUnhandledError, "DCapMov:ExtornaMovimiento Method")
End Function

Public Function ActualizaArendirSust(ByVal pnMovNroSol As Long) As Integer
    On Error GoTo ActualizaArendirSustErr
    Dim sql As String
    ActualizaArendirSust = 1
    
    sql = "UPDATE MovArendirSust SET nMovNroRend=NULL WHERE nMovNro =" & pnMovNroSol & ""
    dbCmact.Execute sql
    ActualizaArendirSust = 0
    Exit Function
ActualizaArendirSustErr:
    Call RaiseError(MyUnhandledError, "DCapMov:ActualizaArendirSust Method")
End Function


