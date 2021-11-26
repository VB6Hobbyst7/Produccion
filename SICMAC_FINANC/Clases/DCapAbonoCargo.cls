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
Dim sSQL As String

Public Function GetnMovNro(ByVal sMovNro As String) As Long
Dim rsMov As Recordset
sSQL = "Select nMovNro From Mov Where cMovNro = '" & sMovNro & "'"
Set rsMov = New Recordset
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
sFecha = Format$(dFecha, "mm/dd/yyyy") & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
sSQL = "Insert Producto (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc) " _
    & "Values ('" & sCuenta & "'," & nTasa & "," & nSaldo & "," & gCapEstActiva & ",'" & sFecha & "',0)"
dbCmact.Execute sSQL
nSaldoDisp = IIf(bCheque, 0, nSaldo)
sFecUltCierre = Format$(DateAdd("d", -1, dFecha), "mm/dd/yyyy")
sSQL = "Insert Captaciones (cCtaCod,nSaldoDisp,nPersoneria,nFirmas,nIntAcum,dUltCierre,dApertura,nPrdCtaTpo,nPrdTasaInteres,cUltimaActualizacion) " _
    & "Values ('" & sCuenta & "'," & nSaldoDisp & "," & nPersoneria & "," & nFirmas & ",0,'" & sFecUltCierre & "','" & sFecha & "'," & nTipoCuenta & "," _
    & nTipoTasa & ",'" & sMovNro & "')"
dbCmact.Execute sSQL
Select Case nProducto
    Case gCapAhorros
        sOrdPag = IIf(bOrdPag, "1", "0")
        sSQL = "Insert CaptacAhorros (cCtaCod,nSaldoAnterior,bOrdPag,nSobregiro,dUltContacto) " _
            & "Values ('" & sCuenta & "',0," & sOrdPag & ",0,'" & sFecha & "')"
    Case gCapPlazoFijo
        sCtaAbonoIntPF = IIf(bCtaAboInt, "1", "0")
        sSQL = "Insert CaptacPlazoFijo (cCtaCod,nPlazo,nIntPag,dRenovacion,nApertura,dAuxiliar,nFormaRetiro,nDuplicado,bAbonoIntCtaAho) " _
            & "Values ('" & sCuenta & "'," & nPlazo & ",0,'" & sFecha & "'," & nSaldo & ",'" & sFecha & "'," & nFormaRetiro & ",0," & sCtaAbonoIntPF & ")"
    Case gCapCTS
        nSaldRetCTS = Round(nSaldo * nPorcRetCTS / 100, 2)
        sSQL = "Insert CaptacCTS (cCtaCod,nSaldRetiro,nIntSaldo,cCodInst) " _
            & "Values ('" & sCuenta & "'," & nSaldRetCTS & ",0,'" & sInstitucion & "')"
End Select
dbCmact.Execute sSQL
If nProducto = gCapPlazoFijo And bCtaAboInt Then
    sSQL = "Insert CaptacCtaAboIntPF (cCtaCod,cCtaCodAbono) " _
        & "Values ('" & sCuenta & "','" & sCtaCodAbono & "')"
    dbCmact.Execute sSQL
End If
AgregaNuevaCaptacion = sCuenta
End Function

Public Sub AgregaNuevoProdPers(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona)
sSQL = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " _
    & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ")"
dbCmact.Execute sSQL
End Sub

Public Function GetDatosOrdenPago(ByVal sCuenta As String, ByVal sNroDoc As String) As Recordset
Dim rsOP As Recordset
sSQL = "SELECT E.cMovNro, D.nTpoDoc, D.cNroDoc, D.cCtaCod, D.nMonto, P.cPersNombre, E.cEstado " _
    & "D.cIFCodPers FROM DocRecOPEst E INNER JOIN DocRecOP D INNER JOIN " & sDBPersona & "Persona P " _
    & "ON D.cIFCodPers = P.cPersCod ON E.nTpoDoc = D.nTpoDoc AND E.cNroDoc = D.cNroDoc AND E.cCtaCod " _
    & "= D.cCtaCod WHERE E.cMovNro IN (SELECT MAX(E1.cMovNro) FROM DocRecOPEst E WHERE E1.nTpoDoc = " _
    & "E.nTpoDoc AND E1.cNroDoc = E.cNroDoc AND E1.cCtaCod = E.cCtaCod)"
Set rsOP = New Recordset
rsOP.CursorLocation = adUseClient
rsOP.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsOP.ActiveConnection = Nothing
Set GetDatosOrdenPago = rsOP
Set rsOP = Nothing
End Function

Public Sub AgregaOrdenPagoEstado(ByVal sCuenta As String, ByVal sNroDoc As String, _
        ByVal sMovNro As String, ByVal nMonto As Double, ByVal nEstadoOP As CaptacOrdPagoEstado)
sSQL = "UPDATE DocRecOP Set nMonto = " & nMonto & " WHERE nTpoDoc = " & TpoDocOrdenPago & " " _
    & "AND cNroDoc = '" & sNroDoc & "' And cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSQL
sSQL = "INSERT DocRecOPEst (nTpoDoc,cNroDoc,cCtaCod,cMovNro,nMonto,nEstado) " _
    & "VALUES (" & TpoDocOrdenPago & ",'" & sNroDoc & "','" & sCuenta & "','" & sMovNro & "'," & nMonto & "," & nEstadoOP & ")"
dbCmact.Execute sSQL
End Sub

Public Sub AgregaCuentaDocumento(ByVal sCuenta As String, ByVal nTpoDoc As TpoDoc, _
        ByVal sNroDoc As String, ByVal sCodIF As String, ByVal sMovNro As String, _
        ByVal nMovNro As Long, Optional nMonto As Double = 0, _
        Optional nEstadoOP As CaptacOrdPagoEstado = gCapOPEstCobrada)
Dim sFecha As String
sFecha = Mid(sMovNro, 1, 4) & "/" & Mid(sMovNro, 5, 2) & "/" & Mid(sMovNro, 7, 2) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
Select Case nTpoDoc
    Case TpoDocCheque
        sSQL = "INSERT DocRecCapta (nTpoDoc,cNroDoc,cPersCod,cIFTpo,cCtaCod,nMonto) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "','" & sCuenta & "'," & nMonto & ")"
    Case TpoDocOrdenPago
        sSQL = "INSERT DocRecOP (nTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers,nEstado,cMovNro) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCuenta & "'," & nMonto & ",'" & sCodIF & "'," & nEstadoOP & ",'" & sMovNro & "')"
    Case TpoDocNotaAbono, TpoDocNotaCargo
End Select
dbCmact.Execute sSQL
sSQL = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
    & "VALUES (" & nMovNro & "," & nTpoDoc & ",'" & sNroDoc & "','" & sFecha & "')"
dbCmact.Execute sSQL
End Sub

Public Sub AgregaMov(ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
        ByVal sGlosa As String, Optional nMovEstado As MovEstado = gMovEstContabMovContable, _
        Optional nMovFlag As MovFlag = gMovFlagVigente)
sSQL = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & nOperacion & "','" & sGlosa & "'," & nMovEstado & "," & nMovFlag & ")"
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

Public Sub UltimaActualizacionCuenta(ByVal sCuenta As String, ByVal sMovNro As String)
'Actualiza la fecha de ultima actualizacion de la cuenta de captaciones
sSQL = "Update Captaciones Set cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSQL
End Sub

Public Sub ActualizaEstadoCuenta(ByVal sCuenta As String, ByVal nEstado As CaptacEstado, _
        ByVal dFecha As Date, ByVal sMovNro As String)
'Actualiza el ultimo estado a la tabla de producto
sSQL = "Update Producto Set nPrdEstado = " & nEstado & ", " _
    & "dPrdEstado = '" & Format$(dFecha, "mm/dd/yyyy") & "' WHERE " _
    & "cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSQL
'Agrega un registro mas de estado a la historia de estados de la cuenta
sSQL = "INSERT CaptacEstado (cCtaCod,nPrdEstado,cMovNro) " _
    & "VALUES ('" & sCuenta & "'," & nEstado & ",'" & sMovNro & "')"
dbCmact.Execute sSQL
End Sub

Public Sub ActualizaEstadoDocRecEst(ByVal nTipoDoc As TpoDoc, ByVal sNroDoc As String, _
        ByVal sCodIF As String, ByVal sMovNro As String, ByVal nEstado As ChequeEstado)
'Actualiza el ultimo estado a la tabla de Documento Recibidos
sSQL = "Update DocRec Set nEstado = " & nEstado & " WHERE " _
    & "nTpoDoc = '" & nTipoDoc & "' AND cNroDoc = '" & sNroDoc & "' AND cPersCod = '" & sCodIF & "'"
dbCmact.Execute sSQL
'Agrega un registro mas de estado a la historia de estados del documento recibido
sSQL = "INSERT DocRecEst (nTpoDoc,cNroDoc,cPersCod,cMovNro,nEstado) " _
    & "VALUES ('" & nTipoDoc & "','" & sNroDoc & "','" & sCodIF & "','" & sMovNro & "'," & nEstado & ")"
dbCmact.Execute sSQL
End Sub

Public Function GetMovExtorno(ByVal sDatoBus As String, Optional nTipoBus As Integer = 0) As Recordset
Dim rsMov As Recordset
Dim sWhere As String
If nTipoBus = 0 Then
    sWhere = "M.cMovNro LIKE '%" & sDatoBus & "%'"
ElseIf nTipoBus = 1 Then
    sWhere = "M.cMovNro LIKE '" & Format$(gdFecSis, "yyyymmdd") & "%' And C.cCtaCod = '" & sDatoBus & "'"
End If
sSQL = "Select M.cMovNro, O.cOpeDesc, C.cCtaCod, ISNULL(MD.nDocTpo,'') nDocTpo, ISNULL(MD.cDocNro,'') cDocNro, " _
    & "M.cMovDesc, C.cOpeCod, C.nMonto, M.nMovNro FROM MovDoc MD RIGHT JOIN Mov M INNER JOIN MovCap C " _
    & "INNER JOIN " & sDBComunes & "OpeTpo O ON C.cOpeCod = O.cOpeCod ON M.nMovNro = C.nMovNro " _
    & "ON MD.nMovNro = M.nMovNro WHERE " & sWhere & " AND C.cOpeCod NOT IN ('" & gAhoEstInacAct & "') " _
    & "AND M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagEliminado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ")" _
    & "AND M.nMovEstado IN (" & gMovEstContabMovContable & ") " _
    & "ORDER BY M.nMovNro"
Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetMovExtorno = rsMov
End Function

Public Sub ActualizaEstadoMov(ByVal nMovNro As Long, ByVal nMovFlag As MovFlag)
sSQL = "Update Mov Set nMovFlag = " & nMovFlag & " Where nMovNro = " & nMovNro
dbCmact.Execute sSQL
End Sub

Public Function EsOrdenPagoEmitida(ByVal sCuenta As String, ByVal nNumOP As Long) As Boolean
Dim rsOP As Recordset
Set rsOP = New Recordset
rsOP.CursorLocation = adUseClient
sSQL = "Select OP.cCtaCod, OP.nInicio, OP.nFin FROM Mov M INNER JOIN MovDocEmitidoRango OP ON " _
    & "M.nMovNro = OP.nMovNro WHERE M.nMovFlag <> " & gMovFlagExtornado & " AND OP.cCtaCod = '" _
    & sCuenta & "' And " & nNumOP & " Between OP.nInicio And OP.nFin"
rsOP.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
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
Dim ClsIni As DClassIni
Set ClsIni = New DClassIni
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
       sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & psCodAge & "','" & psUser & "'"
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
        & "VALUES (" & pnMovNro & "," & pnDocTpo & ",'" & psDocNro & "','" & Format(psDocFecha, "mm/dd/yyyy") & "')"
        
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


Private Function Abono(psCodCta As String, pnMonto As Currency, pgsACDepNA As String, pgsCodAge As String, pgsCodUser As String, NroDoc As String, psDes As String, poCon As DConecta, pgdFecSis As Date)
    Dim sqlA As String
    Dim sqlT As String
    Dim sqlTra As String
    Dim sqlVerifica As String
    Dim sqlNota As String
    Dim lnInteres As Currency
    Dim RegCuenta As ADODB.Recordset
    Set RegCuenta = New ADODB.Recordset
    
    Dim lnSaldoDisponible As Currency
    Dim lnSaldCnt As Currency
    Dim lnSaldDisp As Currency
    Dim lbSAux As Boolean
    Dim lcEst As String
    Dim lNTInt As Currency
    Dim lsHoy As String
    Dim lsAyer As String
    Dim I As Long
    
    lsHoy = Format(pgdFecSis, "mm/dd/yyyy") + " " + Format(Time, "hh:mm:ss AMPM")
    lsAyer = Format(pgdFecSis - 1, "mm/dd/yyyy") + " " + Format(Time, "hh:mm:ss AMPM")

    
    If Mid(psCodCta, 3, 3) = gCapAhorros Then  'Ahorro
        sqlA = "Select * from AhorroC where cCodCta = '" & psCodCta & "'"
        Set RegCuenta = poCon.CargaRecordSet(sqlA)
        
        lnInteres = RegCuenta!nInteres + GetInteres(RegCuenta!nTasaIntAC, RegCuenta!nsalddispac, DateDiff("d", RegCuenta!dUltMovAC, pgdFecSis) - 1, True)
        
        lnSaldoDisponible = RegCuenta!nsalddispac
        lnSaldCnt = RegCuenta!nSaldCntAC + CCur(pnMonto)
        lnSaldDisp = RegCuenta!nsalddispac + CCur(pnMonto)

        lbSAux = False
        If RegCuenta!bInactiva Then
            lcEst = "0"
            lbSAux = True
            'Realiza la Transaccion de activación

            sqlT = "Insert TranDiaria (dFecTran, cCodUsu, cCodOpe, cCodCta, cNumDoc, nMonTran, nSaldCnt, cCodUsuRem, cCodAge, cFlag, nTipCambio) values('" & lsHoy & "' , '" & pgsCodUser & "' , '" _
                    & gsACInaAct & "' , '" & psCodCta & "' , null , " & Str(lnSaldoDisponible) & " , " & lnSaldCnt & ", null , '" & pgsCodAge & "' , NULL,0)"
        Else
            lcEst = "0"
        End If

        If RegCuenta!cTipTasAC = "T" Then
            lNTInt = GetNuevaTasaAC(psCodCta, lnSaldDisp, RegCuenta!nTasaIntAC, poCon)
        Else
            lNTInt = RegCuenta!nTasaIntAC
        End If

        sqlA = " Update AhorroC " _
             & " Set nSaldAntAC=nSaldDispAC, nSaldDispAC=nSaldDispAC+" & pnMonto & ", nSaldCntAC=nSaldCntAC+" & pnMonto & ", dUltMovAC='" & lsAyer & "', nInteres = " & lnInteres & " , dUltActAC='" & lsHoy & "', dUltCntAC = '" + lsHoy + "' ,cCodUsu= '" & pgsCodUser & "', nNumExt = nNumExt + 1 , ntasaintAC = " + Str(lNTInt) + ", bInactiva = " & lcEst & " " _
             & " where cCodCta='" & psCodCta & "'"
        
        sqlTra = " Insert TranDiaria (dFecTran, cCodUsu, cCodOpe, cCodCta, cNumDoc, nMonTran, nSaldCnt, cCodUsuRem, cCodAge, cFlag, nTipCambio) values('" & lsHoy & "' , '" & pgsCodUser & "' , '" _
               & pgsACDepNA & "' , '" & psCodCta & "' , '" & NroDoc & "', " & pnMonto & " , " & Str(lnSaldCnt) & " , null , '" & pgsCodAge & "' , NULL,0)"
        
        sqlVerifica = "Select nNumTran from TranDiaria Where dFecTran = '" & lsHoy & "' And cCodCta = '" & psCodCta & "' AND  cCodAge = '" & pgsCodAge & "' AND cCodUsu = '" & pgsCodUser & "' AND cCodOpe = '" & pgsACDepNA & "' "
        
        RegCuenta.Close
        Set RegCuenta = Nothing
            
        Randomize
        For I = 0 To Rnd(2000) * 1000
        Next I
        
        poCon.Ejecutar sqlA
        If lbSAux Then poCon.Ejecutar sqlT
        poCon.Ejecutar sqlTra
        Set RegCuenta = poCon.CargaRecordSet(sqlVerifica)
        
        If RegCuenta.EOF And RegCuenta.BOF Then
            Set RegCuenta = Nothing
            GoTo ERROR
        End If
                  
        sqlT = Str(RegCuenta!nNumTran)
        
        RegCuenta.Close
        Set RegCuenta = Nothing
                
        sqlNota = " INSERT OtrasOpe(dFecTran,nNumTran,cCodOpe,cGlosa,cNumDoc,cCodPers,cFlag) " _
                & " Values('" & lsHoy & "'," & sqlT & ",'" & pgsACDepNA & "','" & psDes & "','" & NroDoc & "',null,null)"
        poCon.Ejecutar sqlNota
                  
        gbOpeOk = True
        Abono = lnSaldDisp
        
    End If 'fin Ahorro
    
    Exit Function
ERROR:
    
End Function

'***********************************************
' GetInteres: Funcion que devuelve el interes de la cuenta este interes puede ser
' simple o compuesto de acuerdo al ultimo argumento de la funcion
' si este es True es I simple de lo contrario sera compuesto
'***********************************************
Public Function GetInteres(pnTasa As Currency, pnSaldo As Currency, pnDias As Integer, Simple As Boolean) As Double
    Dim lnTasaEfec As Double
    
    If Simple Then
        If pnDias > 0 Then
            lnTasaEfec = (pnTasa / 100) / 360
            GetInteres = Round((lnTasaEfec) * pnDias * pnSaldo, 2)
        Else
            GetInteres = 0
        End If
    Else
        If pnDias > 0 Then
            lnTasaEfec = (pnTasa / 100) / 360
            GetInteres = Round((((lnTasaEfec + 1) ^ pnDias) - 1) * pnSaldo, 2)
        Else
            GetInteres = 0
        End If
    End If

End Function

'*************************************************
'GetNuevaTasaAC:Funcion que devuelve la nueva tasa de Interes de la cuenta
'segun el nuevo saldo de la cuenta, solo se usa para
'ahorro.
'recibe como parametro el numero de cuenta
'El nuevo saldo, el interes actual y la conexion
'*************************************************
Public Function GetNuevaTasaAC(pCuenta As String, pSaldo As Currency, pInt As Double, poCon As DConecta) As Double
    Dim lcCodInt As String
    On Error GoTo ERROR
    lcCodInt = Mid(pCuenta, 3, 4)
    Dim RegCuenta As New ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT nTasaInt FROM TASAINT WHERE cTipTas = 'T' AND nMonInf<= " + Str(pSaldo) + " AND nMonSup>=" + Str(pSaldo) + " AND cCodTas = '" + lcCodInt + "'"
    
    Set RegCuenta = poCon.CargaRecordSet(sql)
    
    If RegCuenta.BOF And RegCuenta.EOF Then
        MsgBox "No Existe un Saldo asignable al monto existente en la cuenta, la cuenta permanecera con el interes anterior "
        GetNuevaTasaAC = pInt
        Set RegCuenta = Nothing
    Else
        GetNuevaTasaAC = RegCuenta!nTasaInt
        RegCuenta.Close
        Set RegCuenta = Nothing
    End If
    
    'Set oCon = Nothing
    
    Exit Function
ERROR:
'        MsgBox Err.Description + "Funciones Principales", vbCritical, frmRealizaOpe.caption
        GetNuevaTasaAC = RegCuenta!nTasaInt
End Function


' Procedimeinto que imprime la boleta de Operación
' luego de una operacion de cajero por impresora
Public Function ImprimeBoletaCad(pgdFecSis As Date, ByVal psTit As String, psText As String, _
                            psCodOpe As String, pnMonto As String, _
                            psCliente As String, psCodCta As String, _
                            psNumDoc As String, pnSaldo As Currency, pnInteresA As String, NomDoc As String, _
                            pnNumExt As Long, pnSaldoC As Double, Optional pbOpSaldoC As Boolean = True, _
                            Optional pbSaldoInt As Boolean = True, Optional psNumDias As String = "----", _
                            Optional psNomAgeRem As String = "", Optional psCodUsuRem As String = "", Optional pbCuenta As Boolean = False, Optional psComCmac As String = "XXX", Optional psLin3 As String = "XXX", Optional psTexto As String = "XXX", Optional pbConfirmacion As Boolean = True) As String
    Dim sFecha As String
    Dim sHora As String
    Dim sSep As Integer
    Dim sIni As Integer
    Dim sMonto As String
    Dim sSDisp As String
    Dim sIntAcum As String
    Dim sMax  As Integer
    Dim sAux As Integer
    Dim lsNegritaOn As String
    Dim lsNegritaOff As String
    Dim lsNroExt As String
    Dim lnTope As Integer
    Dim lsSaldoC As String
    Dim lsCadArg As String
    Dim lsInteres As String
    
    Dim lnCliAux As Long
    Dim lsCliAux1 As String
    Dim lsCliAux2 As String
    
    Dim lnChq As Long
    Dim lsChqAux1 As String
    Dim lsChqAux2 As String
    Dim lsNomAge As String
    
    Dim lnNumLinCmac As Integer
    
    Dim lsCadena As String
    
ETIQ:
    
    On Error GoTo ERROR
    
    lnTope = 0 '6 'Tope de lineas en Boleta
    
    lsNegritaOn = oImpresora.gPrnBoldON
    lsNegritaOff = oImpresora.gPrnBoldOFF
    
    lsNroExt = Str(pnNumExt)
    
    lsCadena = lsCadena & oImpresora.gPrnInicializa
    
    lsCadena = lsCadena & oImpresora.gPrnEspaLineaN      'espaciamiento lineas 1/6 pulg.
    lsCadena = lsCadena & oImpresora.gPrnTamPagina22      'Longitud de página a 22 líneas'
    lsCadena = lsCadena & oImpresora.gPrnTamLetra10CPI      'Tamaño 10 cpi
    lsCadena = lsCadena & oImpresora.gPrnTpoLetraRoman         'Tipo de Letra Sans Serif
    lsCadena = lsCadena & oImpresora.gPrnCondensadaOFF   ' cancela condensada
    lsCadena = lsCadena & oImpresora.gPrnBoldOFF   ' desactiva negrita
    
    sSep = 15
    sIni = 1
    sMax = 33
    sAux = 5
    
    
    sFecha = Format$(pgdFecSis, "dd/mm/yyyy")
    sHora = Format$(Time, "hh:mm:ss")
    sMonto = Format$(pnMonto, "#,##0.00")
    sSDisp = Format$(pnSaldo, "#,##0.00")
    lsSaldoC = Format$(pnSaldoC, "#,##0.00")
    
    lsNomAge = gsNomAge
    
    'lsCadena = lscadena &  oImpresora.gPrnSaltoLinea &
    lsCadena = lsCadena & lsNegritaOn  'Activa Negrita
    lsCadena = lsCadena & Space(sIni) & "CMACT - AHORRO" & Space(19 + sSep + sAux) & "CMACT - AHORRO" & oImpresora.gPrnSaltoLinea
    
    If Mid(psCodCta, 6, 1) = 1 Then
        lsCadena = lsCadena & Space(sIni) & Trim(gsNomAge) & "-SOLES" & Space(sAux + sMax - Len(Trim(gsNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Trim(gsNomAge) & "-SOLES" & Space(sAux + sMax - Len(Trim(gsNomAge)) - Len(lsNroExt) - Len("-SOLES")) & lsNroExt & oImpresora.gPrnSaltoLinea
    Else
        lsCadena = lsCadena & Space(sIni) & Trim(gsNomAge) & "-DOLARES" & Space(sAux + sMax - Len(Trim(gsNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Trim(gsNomAge) & "-DOLARES" & Space(sAux + sMax - Len(Trim(gsNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & oImpresora.gPrnSaltoLinea
    End If
    
    If psNomAgeRem = "" Then
        lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
    Else
        lsCadena = lsCadena & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Space(sAux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))) & "Ag.Rem: " & Trim(psNomAgeRem) & oImpresora.gPrnSaltoLinea
    End If
    
    If psComCmac = "XXX" Then
        If psLin3 = "XXX" Then
            lsCadena = lsCadena & lsNegritaOff & oImpresora.gPrnSaltoLinea    'Desactiva Negrita
        Else
            lsCadena = lsCadena & Space(sIni) & psLin3 & Space(sAux + sSep + sMax - Len(psLin3)) & psLin3 & lsNegritaOff & oImpresora.gPrnSaltoLinea     'Desactiva Negrita
            lnNumLinCmac = 1
        End If
        lnNumLinCmac = 0
    Else
        lsCadena = lsCadena & Space(sIni) & "NroDocCmac:" & psComCmac & Space(sAux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & "NroDocCmac:" & psComCmac & lsNegritaOff & oImpresora.gPrnSaltoLinea     'Desactiva Negrita
        lnNumLinCmac = 1
    End If
    
    lsCadena = lsCadena & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(sAux + sSep - 6) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & oImpresora.gPrnSaltoLinea
    
    'psCliente = PstaNombre(psCliente)
    
    lnCliAux = InStr(1, psCliente, "*", vbTextCompare)
    
    If lnCliAux = 0 Then
        If sAux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + sAux)
        lsCadena = lsCadena & Space(sIni) & ImpreCarEsp(psCliente) & Space(sAux + sMax + sSep - Len(psCliente)) & ImpreCarEsp(psCliente) & oImpresora.gPrnSaltoLinea
    Else
        lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
        lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
        
        If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + sAux)
        If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + sAux)
        
        lsCadena = lsCadena & Space(sIni) & ImpreCarEsp(lsCliAux1) & Space(sAux + sMax + sSep - Len(lsCliAux1)) & ImpreCarEsp(lsCliAux1) & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & Space(sIni) & ImpreCarEsp(lsCliAux2) & Space(sAux + sMax + sSep - Len(lsCliAux2)) & ImpreCarEsp(lsCliAux2) & oImpresora.gPrnSaltoLinea
        
        lnCliAux = 1
    End If
    
    If pbSaldoInt Or pbCuenta Then
        lsCadena = lsCadena & Space(sIni) & "Cuenta:" & psCodCta & Space(14 + sSep + sAux) & "Cuenta:" & psCodCta & oImpresora.gPrnSaltoLinea
    Else
        lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
    End If
    
    psTit = Trim(psTit)
    psTit = CentrarCadena(psTit, 28)
    lsCadena = lsCadena & lsNegritaOn    'Activa Negrita
    lsCadena = lsCadena & Space(sIni + 1) & "-----" & psTit & "-----" & Space(-1 + sSep) & "-----" & psTit & "-----" & oImpresora.gPrnSaltoLinea
    
    
    lnChq = InStr(1, psText, "*", vbTextCompare)
    
    If psTexto = "XXX" Then
        If lnChq = 0 Then
            lsCadena = lsCadena & Space(sIni) & ImpreCarEsp(Trim(psText)) & Space(sMax + 6 - Len(Trim(psText)) - Len(sMonto)) & sMonto & Space(-1 + sSep) & ImpreCarEsp(Trim(psText)) & Space(sMax + 6 - Len(Trim(psText)) - Len(sMonto)) & sMonto & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
        Else
            lsChqAux1 = (Mid(psText, 1, lnChq - 1))
            lsChqAux2 = (Mid(psText, lnChq + 1))
            lsCadena = lsCadena & Space(sIni) & ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Space(-1 + sSep) & ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & Space(sIni) & ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Space(-1 + sSep) & ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & oImpresora.gPrnSaltoLinea
        End If
    Else
        lsCadena = lsCadena & Space(sIni) & ImpreCarEsp(Trim(psTexto)) & Space(sAux + sSep + sMax - Len(Trim(psTexto))) & ImpreCarEsp(Trim(psTexto)) & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
    End If
    
    lsCadena = lsCadena & lsNegritaOff    'Desactiva Negrita
    
    If pbSaldoInt Then
        If pbConfirmacion Then
            If MsgBox("Desea Imprimir el Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                pbConfirmacion = True
            Else
                pbConfirmacion = False
            End If
        Else
            pbConfirmacion = True
        End If
        If pbConfirmacion Then
            lsCadena = lsCadena & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Space(-1 + sSep) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & oImpresora.gPrnSaltoLinea
            If pbOpSaldoC Then
                lsCadena = lsCadena & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Space(-1 + sSep) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & oImpresora.gPrnSaltoLinea
            Else
                lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
            End If
        Else
            lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
            pbSaldoInt = False
        End If
    Else
        lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
    End If
    
    lsInteres = pnInteresA
    
    If pbSaldoInt Then
        If lsInteres <> "No Valido" Then
            lsInteres = Format(lsInteres, "#,##0.00")
            lsCadena = lsCadena & Space(sIni) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Space(-1 + sSep) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & oImpresora.gPrnSaltoLinea
        End If
    Else
        lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
    End If
    
    If Not psNumDoc = "" Then
        lsCadena = lsCadena & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Space(-1 + sSep) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & oImpresora.gPrnSaltoLinea
    Else
        lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
    End If
    
    If Not psNumDias = "----" Then
        lsCadena = lsCadena & Space(sIni) & "Nro Dias Interes" & Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)) & psNumDias & Space(-1 + sSep) & "Nro Dias Interes" & Space(sMax + 6 - Len(psNumDias) - Len("Nro Dias Interes")) & psNumDias & oImpresora.gPrnSaltoLinea
        lnTope = 4 - lnCliAux
    Else
        lnTope = 3 - lnCliAux
    End If
    
    lsCadena = lsCadena & Space(sIni) & "---------------------------------------" & Space(-1 + sSep) & "---------------------------------------" & oImpresora.gPrnSaltoLinea
    If psCodUsuRem = "" Then
        lsCadena = lsCadena & Space(sIni) & ImpreCarEsp(gsCodUser) & Space(29 + sSep + sAux) & ImpreCarEsp(gsCodUser) & oImpresora.gPrnSaltoLinea
    Else
        lsCadena = lsCadena & Space(sIni) & "Loc/Rem" & Space(sMax + sAux - Len("Loc/Rem") - 1 - 8) & ImpreCarEsp(gsCodUser) & "/" & ImpreCarEsp(psCodUsuRem) & Space(sSep) & "Loc/Rem" & Space(sMax + sAux - Len("Loc/Rem") - 1 - 8) & ImpreCarEsp(gsCodUser) & "/" & ImpreCarEsp(psCodUsuRem) & oImpresora.gPrnSaltoLinea
    End If
    
    'lsCadena = lscadena &  oImpresora.gPrnSaltoPagina
    For sAux = 1 To (lnTope - lnNumLinCmac)
        lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
    Next sAux
    
    ImprimeBoletaCad = lsCadena
    
    Exit Function
ERROR:
        Close nFicSal
        If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
            GoTo ETIQ
        End If
End Function


Function Cargo(psCodCta As String, pnMonto As Currency, pgsACRetNC As String, pgsCodAge As String, pgsCodUser As String, NroDoc As String, psDes As String, poCon As DConecta, pgdFecSis As Date) As Currency
    Dim sqlA As String
    Dim sqlT As String
    Dim sqlTra As String
    Dim sqlVerifica As String
    Dim sqlNota As String
    Dim lnInteres As Currency
    Dim RegCuenta As ADODB.Recordset
    Set RegCuenta = New ADODB.Recordset
    
    Dim lnSaldoDisponible As Currency
    Dim lnSaldCnt As Currency
    Dim lnSaldDisp As Currency
    Dim lbSAux As Boolean
    Dim lcEst As String
    Dim lNTInt As Currency
    Dim lsHoy As String
    Dim lsAyer As String
    Dim I As Long
    Dim inicioTrans As Boolean

    lsHoy = Format(pgdFecSis, "mm/dd/yyyy") + " " + Format(Time, "hh:mm:ss AMPM")
    lsAyer = Format(pgdFecSis - 1, "mm/dd/yyyy") + " " + Format(Time, "hh:mm:ss AMPM")
    
    lbSAux = False
    
    If Mid(psCodCta, 3, 3) = gCapAhorros Then 'Ahorro
        sqlA = "Select * from AhorroC where cCodCta = '" & psCodCta & "'"
        Set RegCuenta = poCon.CargaRecordSet(sqlA)
        
        lnInteres = RegCuenta!nInteres + GetInteres(RegCuenta!nTasaIntAC, RegCuenta!nsalddispac, DateDiff("d", RegCuenta!dUltMovAC, pgdFecSis) - 1, True)

        lnSaldCnt = RegCuenta!nSaldCntAC - CCur(pnMonto)
        lnSaldDisp = RegCuenta!nsalddispac - CCur(pnMonto)
        
        If RegCuenta!cTipTasAC = "T" Then
            lNTInt = GetNuevaTasaAC(psCodCta, lnSaldDisp, RegCuenta!nTasaIntAC, poCon)
        Else
            lNTInt = RegCuenta!nTasaIntAC
        End If
    
        lbSAux = False
        If RegCuenta!bInactiva Then
            lcEst = "0"
            lbSAux = True
            'Realiza la Transaccion de activación
            sqlT = " Insert TranDiaria (dFecTran, cCodUsu, cCodOpe, cCodCta, cNumDoc, nMonTran, nSaldCnt, cCodUsuRem, cCodAge, cFlag, nTipCambio) values('" & lsHoy & "' , '" & pgsCodUser & "' , '" _
                 & gsACInaAct & "' , '" & psCodCta & "' , null , " & Str(Round(lnSaldDisp, 2)) & " ," & Str(lnSaldCnt) & ", null , '" & pgsCodAge & "' , NULL,0)"
        Else
            lcEst = "0"
        End If

        'Actualiza los datos de la tabla de Ahorro
        sqlA = " Update AhorroC " _
             & " set nSaldAntAC=nSaldDispAC, nSaldDispAC=nSaldDispAC-" & Str(Round(CCur(pnMonto), 2)) & ", dUltCntAC = '" + lsHoy + "', bInactiva = " & lcEst & "  ," _
             & " nSaldCntAC=nSaldCntAC-" & Str(Round(CCur(pnMonto), 2)) & ", dUltMovAC='" & lsAyer & "', ntasaintAC = " + Str(Round(lNTInt, 2)) + " " _
             & " , dUltActAC='" & lsHoy & "', cCodUsu='" & gsCodUser & "', nNumExt = nNumExt + 1 ,nInteres = " + Str(Round(lnInteres, 2)) _
             & " where cCodCta='" & psCodCta & "'"

        sqlTra = "Insert TranDiaria (dFecTran, cCodUsu, cCodOpe, cCodCta, cNumDoc, nMonTran, nSaldCnt, cCodUsuRem, cCodAge, cFlag, nTipCambio) values('" & lsHoy & "' , '" _
               & pgsCodUser & "', '" & pgsACRetNC & "','" & psCodCta & "','" & NroDoc & "', " _
               & "-" + Str(Round(CCur(pnMonto), 2)) & ", " & Str(Round(lnSaldCnt, 2)) & ",null,'" & pgsCodAge & "', Null,0)"
    
        sqlVerifica = " Select nNumTran from TranDiaria Where dFecTran = '" & lsHoy & "' And cCodCta = '" & psCodCta & "' AND  cCodAge = '" & pgsCodAge & "' AND cCodUsu = '" & pgsCodUser & "' and cCodOpe = '" & pgsACRetNC & "'"
        
        'Inserta nueva operación en Transacción Diaria
        RegCuenta.Close
        Set RegCuenta = Nothing

        Randomize
        For I = 0 To Rnd(2000) * 1000
        Next I
        
        If Not TieneSaldo(psCodCta, Str(pnMonto), poCon) Then
            MsgBox "Se ha realizado otra operación de Retiro, la cuenta ya no tiene saldo suficiente", vbInformation, "Aviso"
            gbOpeOk = False
            Cargo = 0
            Exit Function
        End If

        poCon.Ejecutar sqlA
        poCon.Ejecutar sqlTra
        If lbSAux Then poCon.Ejecutar sqlT
        
        Set RegCuenta = poCon.CargaRecordSet(sqlVerifica)
        
        If RegCuenta.EOF And RegCuenta.BOF Then
            Set RegCuenta = Nothing
            GoTo ERROR
        End If
          
        sqlA = Str(RegCuenta!nNumTran)
        
        RegCuenta.Close
        Set RegCuenta = Nothing
                
        sqlNota = " INSERT OtrasOpe(dFecTran,nNumTran,cCodOpe,cGlosa,cNumDoc,cCodPers,cFlag) " _
                & " values('" & lsHoy & "'," & sqlA & ",'" & pgsACRetNC & "','" & psDes & "','" & NroDoc & "',null,null)"

        poCon.Ejecutar sqlNota
 
        gbOpeOk = True
        Cargo = lnSaldDisp
    
    End If 'fin Ahorro
    Exit Function
ERROR:
    
End Function


'*************************************************
'TieneSaldo: Funcion que devuelve el true si la cuenta tiene
'el saldo suficiente para realizar el retiro que se esta realizando
'se le ingresa como paramentro el Nro de cuenta el monto a retirar y
'conexion a la agencia que es propietaria de esta cuenta
'con que se esta haciendo un retiro con orden de pago
'caso contrario es False
'*************************************************
Public Function TieneSaldo(psCodCta As String, pnMonto As String, poCon As DConecta) As Boolean
    Dim RegCuenta As New ADODB.Recordset
    Dim lnMontoMinimo As Double
    Dim lnCuentaDescarAuto As Double
    Dim lnResta As Double
    Dim sql As String
    On Error GoTo ERROR
        
    Select Case Mid(psCodCta, 3, 3)
        Case gCapAhorros
                sql = "Select nSaldDispAC as nSaldDisp from AhorroC where cCodCta = '" & psCodCta & "' and cEstCtaAC Not in ('C','U') "
        Case gCapPlazoFijo
                sql = "Select nIntDev as nSaldDisp from PlazoFijo where cCodCta = '" & psCodCta & "' and cEstCtaPF Not in ('C','U')"
        Case gCapCTS
                sql = "Select nSaldDispCTS as nSaldDisp from CTS where cCodCta = '" & psCodCta & "' and cEstCtaCTS Not in ('C','U')"
    End Select
    
    Set RegCuenta = poCon.CargaRecordSet(sql)
    
    If RSVacio(RegCuenta) Then
        Set RegCuenta = Nothing
        TieneSaldo = False
        Exit Function
    End If
    
    Select Case Mid(psCodCta, 3, 3)
        Case gCapAhorros
            If Mid(psCodCta, 6, 1) = "1" Then
                lnMontoMinimo = ReadParametros("23102", poCon)
                lnResta = lnMontoMinimo
                If RegCuenta!nSaldDisp - lnResta >= CCur(pnMonto) Then
                    RegCuenta.Close
                    Set RegCuenta = Nothing
                    TieneSaldo = True
                Else
                    RegCuenta.Close
                    Set RegCuenta = Nothing
                    TieneSaldo = False
                End If
            Else
                lnMontoMinimo = ReadParametros("23103", poCon)
               
                lnResta = lnMontoMinimo
                If RegCuenta!nSaldDisp - lnResta >= CCur(pnMonto) Then
                    RegCuenta.Close
                    Set RegCuenta = Nothing
                    TieneSaldo = True
                Else
                    RegCuenta.Close
                    Set RegCuenta = Nothing
                    TieneSaldo = False
                End If
            End If
        
        Case gCapPlazoFijo
                If RegCuenta!nSaldDisp >= CCur(pnMonto) Then
                RegCuenta.Close
                Set RegCuenta = Nothing
                TieneSaldo = True
            Else
                RegCuenta.Close
                Set RegCuenta = Nothing
                TieneSaldo = False
            End If
        Case gCapCTS
            
            If (RegCuenta!nSaldDisp) >= CCur(pnMonto) Then
                RegCuenta.Close
                Set RegCuenta = Nothing
                TieneSaldo = True
            Else
                RegCuenta.Close
                Set RegCuenta = Nothing
                TieneSaldo = False
            End If
    End Select

    Exit Function
ERROR:
    MsgBox Err.Description & ". Funciones Principales", vbCritical, "Aviso"
End Function

'****************************************
'ReadParametros: Funcion que devuelve el valor de un parametro de
'la tabla parametros, se ingresa el codigo del parametro
'****************************************
Public Function ReadParametros(txtCodPar As String, poCon As DConecta) As Double

    Dim RecVar As New ADODB.Recordset
    Dim qryVar As String
    
    On Error GoTo ERROR
    
    qryVar = "SELECT nValor1 FROM Parametro WHERE cCodPar = '" & txtCodPar & "'"
    Set RecVar = poCon.CargaRecordSet(qryVar)
    
    If RecVar.EOF And RecVar.BOF Then
        ReadParametros = 0
        Set RecVar = Nothing
    Else
        ReadParametros = RecVar!nValor1
        RecVar.Close
        Set RecVar = Nothing
    End If
    
    Exit Function
ERROR:
    MsgBox "Error en Conexión + " + Err.Description, vbCritical, "Aviso"
    ReadParametros = -1
End Function


