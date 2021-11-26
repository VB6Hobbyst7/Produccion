VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMDocRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsServerCom As String
Dim vsServerPers As String

Public Function RegistroChequesNegocio(ByVal psMovNro As String, ByVal psOpeCod As String, _
            ByVal psMovDesc As String, ByVal psTipoIF As String, _
            ByVal psNumCheque As String, ByVal psPersCodIF As String, ByVal pnPlaza As ChequePlaza, _
            ByVal psCtaCheque As String, ByVal pnImporte As Double, ByVal pdFechaReg As Date, _
            ByVal pdFechaVal As Date, ByVal psFormatoFecha As String, ByVal pnMonedaCheque As Moneda, _
            Optional ByVal pnEstadoCheque As ChequeEstado = gChqEstEnValorizacion, _
            Optional ByVal pnEstadoCG As CGEstadosChq = gCGEstadosChqRecibido, _
            Optional ByVal pnConfirmaCaja As CGEstadoConfCheque = ChqCGSinConfirmacion, _
            Optional psAreaCod As String, Optional psAgeCod As String, _
            Optional pnProducto As Producto = 0, Optional ByVal psPersCodIF1 As String = "", Optional ByVal psNumChequeComp As String = "") As Integer
 'Se agrego parametro psPersCodIF1 = girador
 'MADM 20110701
Dim oMov As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim lnMovNro As Long

On Error GoTo RegistroChequesNegocioErr
Set oMov = New COMDMov.DCOMMov
'oMov.inicio psFormatoFecha
oMov.BeginTrans
lbTrans = True
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)
If psNumCheque <> "" Then
    oMov.InsertaMovDoc lnMovNro, TpoDocCheque, psNumCheque, Format$(pdFechaReg, "mm/dd/yyyy")
End If
'grabacion dentro de la tabla de cheques
'MADM 20110224 psPersCodIF1 = girador
oMov.InsertaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pnPlaza, psCtaCheque, pnImporte, pdFechaVal, _
                        pdFechaVal, pnEstadoCG, pnConfirmaCaja, pnMonedaCheque, gChqEstRegistrado, psAreaCod, psAgeCod, pnProducto, psPersCodIF1, psNumChequeComp
                        
oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pdFechaReg, gChqEstRegistrado, psMovNro, psCtaCheque



RegistroChequesNegocio = 0
oMov.CommitTrans
lbTrans = False
Exit Function
RegistroChequesNegocioErr:
    If lbTrans Then
        lbTrans = False
        oMov.RollbackTrans
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function RegistroChequesContab(ByVal psMovNro As String, ByVal psOpeCod As String, _
                                      ByVal psMovDesc As String, ByVal pnMovRef As Long, _
                                      ByVal psCtaDebe As String, ByVal psObjCodProd As String, ByVal psAreaCod As String, ByVal psAgeCod As String, _
                                      ByVal psCtaHaber As String, ByVal rsObjHaber As adodb.Recordset, _
                                      ByVal psNumCheque As String, ByVal psPersCodIF As String, ByVal psTipoIF As String, ByVal pnPlaza As ChequePlaza, _
                                      ByVal psCtaCheque As String, ByVal pnImporte As Currency, ByVal pdFechaReg As Date, _
                                      ByVal pdFechaVal As Date, ByVal psFormatoFecha As String, ByVal pnMonedaCheque As Moneda, _
                                      Optional ByVal pnEstadoCheque As ChequeEstado = gChqEstEnValorizacion, _
                                      Optional ByVal pnEstadoCG As CGEstadosChq = gCGEstadosChqRecibido, _
                                      Optional ByVal pnConfirmaCaja As CGEstadoConfCheque = ChqCGSinConfirmacion, _
                                      Optional ByVal psAreaCodChq As String, Optional ByVal psAgeCodChq As String, Optional ByVal psPersCodIF1 As String = "", Optional ByVal pValTottxtChqNumCheque As String = "") As Integer

'MADM 20110702 - 20110225 - Girador
Dim oMov As COMDMov.DCOMMov
Dim lsSubCuenta As String
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New COMDMov.DCOMMov

On Error GoTo RegistroChequesContabErr
lsSubCuenta = ""
'oMov.inicio psFormatoFecha
oMov.BeginTrans
lbTrans = True
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc
lnMovNro = oMov.GetnMovNro(psMovNro)
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1
oMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe, pnImporte
lnMovOrden = lnMovOrden + 1
oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, psObjCodProd
lnMovOrden = lnMovOrden + 1
oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
oMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, psAgeCod, psAreaCod

'guardamos la cuenta de ContraResta
lsSubCta = ""
If Not rsObjHaber Is Nothing Then
    If rsObjHaber.State = adStateOpen Then
        If Not rsObjHaber.EOF And Not rsObjHaber.BOF Then
            Do While Not rsObjHaber.EOF
                lsSubCta = lsSubCta + rsObjHaber!SubCta
                rsObjHaber.MoveNext
            Loop
        End If
    End If
End If
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber + lsSubCta, pnImporte * -1
If Not rsObjHaber Is Nothing Then
    If rsObjHaber.State = adStateOpen Then
        rsObjHaber.MoveFirst
        Do While Not rsObjHaber.EOF
            lnMovOrden = lnMovOrden + 1
            Select Case rsObjHaber!cObjetoCod
                Case ObjCMACAgenciaArea
                    oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!cObjetoCod
                    oMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, Mid(rsObjHaber!objeto, 4, 2), Mid(rsObjHaber!objeto, 1, 3)
                Case ObjDescomEfectivo
                    oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!cObjetoCod
                    oMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!objeto
                Case ObjEntidadesFinancieras
                    oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!cObjetoCod
                    oMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(rsObjHaber!objeto, 4, 13), Mid(rsObjHaber!objeto, 1, 2), Mid(rsObjHaber!objeto, 18, 10)
                Case Else
                    oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!objeto
            End Select
            rsObjHaber.MoveNext
        Loop
    End If
End If
If psNumCheque <> "" Then
    oMov.InsertaMovDoc lnMovNro, TpoDocCheque, psNumCheque, pdFechaReg
End If
'grabacion dentro de la tabla de cheques
oMov.InsertaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pnPlaza, psCtaCheque, pnImporte, pdFechaReg, _
                        pdFechaVal, pnEstadoCG, pnConfirmaCaja, pnMonedaCheque, , psAreaCodChq, psAgeCodChq, , psPersCodIF1, pValTottxtChqNumCheque

oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pdFechaReg, gChqEstEnValorizacion, psMovNro
If pnMovRef <> 0 Then
    oMov.InsertaMovRef lnMovNro, pnMovRef
End If
oMov.CommitTrans
lbTrans = False
RegistroChequesContab = 0
Exit Function
RegistroChequesContabErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetDatosCheques(ByVal psNumChq As String, ByVal psPersCodIF As String) As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
Dim oConect As COMConecta.DCOMConecta

Set oConect = New COMConecta.DCOMConecta
Set rs = New adodb.Recordset

If oConect.AbreConexion = False Then Exit Function

sql = "SELECT   DR.CPERSCOD , LEFT(P.CPERSNOMBRE,40) AS IFNOMBRE, " _
    & "         DR.nTpoDoc, DR.cNroDoc, DR.bPlaza , LEFT(C1.cConsDescripcion,30)  as PlazaDesc, " _
    & "         DR.cIFCta, DR.nMonto, DR.dValorizaRef, " _
    & "         DR.dValorizacion, DR.cDepIF, DR.nConfCaja, " _
    & "         DRE.nEstado , Left(C.cConsDescripcion, 30) as cEstado, M.CMOVDESC, M.cMovNro  " _
    & " FROM    DOCREC DR " _
    & "         JOIN DOCRECEST DRE ON   DRE.nTPODOC = DR.nTPODOC " _
    & "                                 AND DRE.CNRODOC=DR.CNRODOC " _
    & "                                 AND DR.cPersCod = DRE.cPersCod " _
    & "         JOIN " & vsServerCom & "CONSTANTE C   ON  C.nConsValor = DRE.nEstado  " _
    & "         JOIN " & vsServerPers & "PERSONA P ON P.CPERSCOD = DR.CPERSCOD " _
    & "         JOIN " & vsServerCom & "CONSTANTE C1  ON C1.nConsValor = DR.bPlaza " _
    & "         JOIN MOV M ON M.cMovNro = DRE.cMovNro " _
    & " WHERE   C.nCONSCOD ='" & gChequeEstado & "' AND C1.nCONSCOD ='" & gChequePlaza & "' AND " _
    & "         DRE.cMovNro = ( Select MAX(cMovNro) " _
    & "                         FROM    DocRecEst D " _
    & "                         Where   DRE.nTPODOC = D.nTPODOC AND DRE.CNRODOC=D.CNRODOC " _
    & "                                 AND DR.cPersCod = D.cPersCod) " _
    & "         AND DR.cNroDoc='" & psNumChq & "' AND DR.cPersCod ='" & psPersCodIF & "'"

Set rs = oConect.CargaRecordSet(sql)
Set GetDatosCheques = rs
oConect.CierraConexion
Set oConect = Nothing
End Function

Private Sub Class_Initialize()
Dim oIni As COMConecta.DCOMClasIni
Set oIni = New COMConecta.DCOMClasIni

vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing

End Sub

Public Function GetChequesValidos(ByVal dFecha As Date, Optional nMoneda As Moneda = gMonedaNacional) As Recordset
Dim sSql As String
Dim clsConecta As COMConecta.DCOMConecta
sSql = "Select cPersNombre, cNroDoc, nMonto = nMonto - nMontoUsado, cPersCod, dValorizaRef, nDiasVal, cIFCta  From ( " _
    & "SELECT P.cPersNombre, D.cNroDoc, D.nMonto, P.cPersCod, dValorizaRef, SUM(ISNULL(DC.nMonto,0)) nMontoUsado,DP.nDiasVal,  " _
    & "D.cIFCta FROM (Select DC.nMonto, DC.nTpoDoc, DC.cNroDoc, DC.cPersCod, DC.cIFTpo From Mov M JOIN DocRecCapta DC " _
    & "ON M.nMovNro = DC.nMovNro Where M.nMovFlag = 0) DC RIGHT JOIN DocRec D INNER JOIN DocRecEst E INNER JOIN Persona P ON E.cPersCod = P.cPersCod " _
    & "ON D.nTpoDoc = E.nTpoDoc AND D.cNroDoc = E.cNroDoc AND D.cPersCod = E.cPersCod AND D.cIFTpo = E.cIFTpo " _
    & "ON DC.nTpoDoc = D.nTpoDoc AND DC.cNroDoc = D.cNroDoc AND DC.cPersCod = D.cPersCod AND DC.cIFTpo = D.cIFTpo " _
    & " JOIN DocRecParametro DP ON  D.bPlaza =DP.nPlaza And DP.nMoneda = " & nMoneda & "" _
    & " WHERE D.nEstado IN (" & gChqEstRegistrado & "," & gChqEstEnValorizacion & "," & gsChqEstExtornado & ") AND E.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' " _
    & " AND E.cMovNro IN (Select MAX(cMovNro) FROM DocRecEst E1 " _
    & " WHERE E1.nTpoDoc = E.nTpoDoc AND E1.cNroDoc = E.cNroDoc AND E1.cPersCod = E.cPersCod AND E1.cIFTpo = E.cIFTpo) " _
    & " And D.nMoneda = " & nMoneda & " Group by P.cPersNombre, D.cNroDoc, D.nMonto, P.cPersCod, dValorizaRef,DP.nDiasVal, D.cIFCta  " _
    & ") A Where nMonto - nMontoUsado > 0"

Set clsConecta = New COMConecta.DCOMConecta
If clsConecta.AbreConexion = False Then Exit Function
Set GetChequesValidos = clsConecta.CargaRecordSet(sSql)
clsConecta.CierraConexion
Set clsConecta = Nothing
End Function
Public Function GetChequesCreditos(ByVal dFecha As Date, Optional nMoneda As Moneda = gMonedaNacional) As Recordset
Dim sSql As String
Dim clsConecta As COMConecta.DCOMConecta

sSql = " SELECT DISTINCT cPersNombre, cNroDoc, nMonto , D.cPersCod, dValorizaRef,nDiasVal FROM DOCREC D  "
sSql = sSql & "  JOIN PERSONA P ON P.CPERSCOD=D.CPERSCOD "
sSql = sSql & "  JOIN MOVDOC MD ON MD.CDOCNRO=D.CNRODOC "
sSql = sSql & "  JOIN DocRecParametro DP ON  D.bPlaza =DP.nPlaza "
sSql = sSql & "  LEFT JOIN (SELECT * FROM MOV WHERE NMOVFLAG=0 AND COPECOD='200252' ) M ON M.NMOVNRO=MD.NMOVNRO "
sSql = sSql & "  WHERE D.NESTADO IN (1,2) AND D.NPRODUCTO NOT IN (233,234,232) AND M.NMOVNRO IS NULL  and D.NMONEDA=" & nMoneda

Set clsConecta = New COMConecta.DCOMConecta
If clsConecta.AbreConexion = False Then Exit Function
Set GetChequesCreditos = clsConecta.CargaRecordSet(sSql)
clsConecta.CierraConexion
Set clsConecta = Nothing

End Function


Public Function GetCheques(ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psAgeCod As String, Optional psNumCheque As String) As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltro As String


Set oCon = New COMConecta.DCOMConecta
lsFiltro = ""
If Trim(psNumCheque) <> "" Then
    lsFiltro = " AND DR.cNroDoc LIKE '%" & psNumCheque & "%'"
End If

If oCon.AbreConexion = False Then Exit Function

sql = " Select  DISTINCT DR.cNroDoc, P.cPersNombre, ISNULL(DRC.cCtaCod,'') AS cCtaCod, DR.nMonto," _
    & "         DR.dValorizaRef, DR.dValorizacion, DR.bPlaza, DR.cIFCta, DR.cDepIF, DR.nConfCaja, " _
    & "         ISNULL(C1.cConsDescripcion,'') AS cMoneda, ISNULL(C.cConsDescripcion,'') AS EstActual, " _
    & "         DR.nMoneda , ISNULL(C2.cConsDescripcion,'') AS cPlaza , DR.cPersCod, DR.cIFTpo , DR.nEstado AS nEstado   " _
    & " From    DocRec DR " _
    & "         JOIN MOVDOC MD ON  MD.nDocTpo= DR.nTpoDoc AND DR.cNroDoc=MD.cDocNro " _
    & "         JOIN MOV    M ON M.NMOVNRO = MD.NMOVNRO " _
    & "         JOIN " & vsServerPers & "PERSONA P              ON P.cPersCod =DR.cPersCod " _
    & "         LEFT JOIN " & vsServerCom & "CONSTANTE C        ON  C.nConsValor= DR.nEstado AND C.nConsCod =" & gChequeEstado & "" _
    & "         LEFT JOIN " & vsServerCom & "CONSTANTE C1       ON  C1.nConsValor= DR.nMoneda AND C1.nConsCod =" & gMoneda & "" _
    & "         LEFT JOIN CONSTANTE C2       ON  C2.nConsValor= DR.bPlaza  AND C2.nConsCod =" & gChequePlaza & "  " _
    & "         LEFT JOIN DocRecCapta DRC   ON  DRC.nTpoDoc=DR.nTpoDoc AND DRC.cNrodoc=DR.cNrodoc " _
    & "                                         AND DRC.cPersCod=DR.cPersCod AND DRC.cIFTpo=DR.cIFTpo " _
    & " WHERE   DR.cDepIF=" & gCGEstadosChqRecibido & "  AND SUBSTRING(M.cMovNro,18,2) ='" & psAgeCod & "' AND M.NMOVFLAG =" & gMovFlagVigente & " " _
    & "         AND SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' and '" & Format(pdHasta, "yyyymmdd") & "' " & lsFiltro

Set rs = oCon.CargaRecordSet(sql)
Set GetCheques = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetPersCuentaAho(ByVal psCodCta As String) As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltro As String


Set oCon = New COMConecta.DCOMConecta
'lsFiltro = ""
'If Trim(psNumCheque) <> "" Then
'    lsFiltro = " AND DR.cNroDoc LIKE '%" & psNumCheque & "%'"
'End If

If oCon.AbreConexion = False Then Exit Function

sql = "     Select  P.cPersNombre as Nombre, T1.cConsDescripcion  as Relacion," _
        & "         p.nPersPersoneria  as Personeria " _
        & " FROM    Persona P " _
        & "         JOIN ProductoPersona PP ON P.cPersCod = PP.cPersCod " _
        & "         JOIN Constante T1 ON PP.NPrdPersRelac = T1.nConsValor " _
        & " WHERE   PP.cCtaCod = '" & psCodCta & "' AND T1.nCONSCOD = " & gCaptacRelacPersona & " " _
        & " ORDER BY Nombre "

Set rs = oCon.CargaRecordSet(sql)
Set GetPersCuentaAho = rs
oCon.CierraConexion
Set oCon = Nothing
End Function


Public Function GetEstadosCheques(ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF) As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function

sql = "      Select  CONVERT(VARCHAR(12),Convert(datetime, Substring(DR.cMovNro,1,8)),103) as Fecha ," _
        & "          ISNULL(C.cConsDescripcion,'') AS EstActual " _
        & "  From    DocRecEst DR " _
        & "          LEFT JOIN CONSTANTE C        ON  C.nConsValor= DR.nEstado AND C.nConsCod =1007 " _
        & "  Where  DR.nTpoDoc = " & TpoDocCheque & " And DR.cNroDoc = '" & psNroDoc & "' " _
        & "         And DR.cPersCod ='" & psPersCod & "' And DR.cIFTpo='" & Format(psIFTpo, "00") & "' " _
        & " ORDER BY DR.cMovNro, DR.nEstado ASC  "

Set rs = oCon.CargaRecordSet(sql)
Set GetEstadosCheques = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GrabaCambioEstadoCheque(psFormatoFecha As String, _
                                    ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psOpeDesc As String, _
                                    ByVal psMovDesc As String, _
                                    ByVal psNumCheque As String, ByVal psPersCodIF As String, _
                                    ByVal psTipoIF As String, _
                                    ByVal pnImporte As Currency, ByVal pnEstadoCheque As ChequeEstado, _
                                    ByVal pnConfirmaCaja As CGEstadoConfCheque, ByVal pdFechaVal As Date, _
                                    ByVal psCodCtaAho As String, ByVal pdFecSis As Date) As Integer
Dim oMov As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim lnMovNro As Long

On Error GoTo GrabaCambioEstadoChequeErr
Set oMov = New COMDMov.DCOMMov
GrabaCambioEstadoCheque = 1
'oMov.inicio psFormatoFecha
oMov.BeginTrans
lbTrans = True
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovOpeVarias lnMovNro, psNumCheque, Left(psOpeDesc, 12), pnImporte
Select Case psOpeCod
    Case gOpeChequesAnulacion, gOpeChequesRechazo
        oMov.ActualizaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, , , , , , , , , pnEstadoCheque
        oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pdFecSis, pnEstadoCheque, psMovNro
    Case gOpeChequesValorización
        oMov.ActualizaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, , , , , , , pnConfirmaCaja, , pnEstadoCheque
        oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pdFecSis, pnEstadoCheque, psMovNro
        'falta actualizar el saldo de la cuenta de ahorros que afecto el cheque
        If psCodCtaAho <> "" Then
        End If
    Case gOpeChequesModFecVal
        oMov.ActualizaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, , pdFechaVal, , , , , , , pnEstadoCheque
End Select
GrabaCambioEstadoCheque = 0
oMov.CommitTrans
lbTrans = False
Exit Function
GrabaCambioEstadoChequeErr:
    If lbTrans Then
        lbTrans = False
        oMov.RollbackTrans
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetOpeCheques(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psAgeCod As String, Optional psNumCheque As String) As adodb.Recordset
Dim sql As String
Dim rs As New adodb.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltro As String


Set oCon = New COMConecta.DCOMConecta
lsFiltro = ""
If Trim(psNumCheque) <> "" Then
    lsFiltro = " AND DR.cNroDoc LIKE '%" & psNumCheque & "%'"
End If

If oCon.AbreConexion = False Then Exit Function

sql = " Select  CONVERT(VARCHAR(12), CONVERT(DATETIME,SUBSTRING(M.cMovNro,1,8)),103) AS FECHA, " _
    & "         RIGHT(M.CMOVNRO,4) AS CUSER, DR.cNroDoc, p.cPersNombre, " _
    & "         ISNULL(C.cConsDescripcion,'') AS EstActual, ISNULL(C1.cConsDescripcion,'') AS cMoneda, " _
    & "         D.nMonto, ISNULL(DRC.cCtaCod,'') AS cCtaCod,  p.cPersCod , M.NMOVNRO , D.cIFTpo " _
    & " From    Mov M " _
    & "         JOIN DocRecEst DR on DR.cMovNro = M.cMovnro " _
    & "         JOIN DocRec D on D.nTpoDoc =DR.nTpoDoc and D.cNroDoc = DR.cNroDoc " _
    & "                         and D.cPersCod = Dr.cPersCod and D.cIFTpo=DR.cIFTpo " _
    & "         JOIN " & vsServerPers & "PERSONA P              ON P.cPersCod =DR.cPersCod " _
    & "         LEFT JOIN " & vsServerCom & "CONSTANTE C        ON  C.nConsValor= DR.nEstado AND C.nConsCod =" & gChequeEstado & "" _
    & "         LEFT JOIN " & vsServerCom & "CONSTANTE C1       ON  C1.nConsValor= D.nMoneda  AND C1.nConsCod =" & gMoneda & "" _
    & "         LEFT JOIN DocRecCapta DRC   ON  DRC.nTpoDoc=D.nTpoDoc AND DRC.cNrodoc=D.cNrodoc " _
    & "                                         AND DRC.cPersCod=D.cPersCod AND DRC.cIFTpo=D.cIFTpo " _
    & " WHERE   M.COPECOD IN ('" & psOpeCod & "') AND M.NMOVFLAG =" & gMovFlagVigente & " " _
    & "         AND SUBSTRING(M.cMovNro,18,2) ='" & psAgeCod & "' " _
    & "         AND SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' and '" & Format(pdHasta, "yyyymmdd") & "' " & lsFiltro & "" _
    & "         AND NOT EXISTS (SELECT  M1.NMOVNRO " _
    & "                         FROM    MOV M1 " _
    & "                                 JOIN MOVREF MR ON MR.NMOVNRO=M1.NMOVNRO  " _
    & "                         WHERE   MR.NMOVNROREF=M.NMOVNRO AND M1.NMOVFLAG =" & gMovFlagVigente & ")"

Set rs = oCon.CargaRecordSet(sql)
Set GetOpeCheques = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GrabaExtornoCheque(ByVal psFormatoFecha As String, ByVal pnMovNroAnt As Long, _
                                   ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psOpeDesc As String, _
                                    ByVal psMovDesc As String, _
                                    ByVal psNumCheque As String, ByVal psPersCodIF As String, _
                                    ByVal psTipoIF As String, _
                                    ByVal pnImporte As Currency, ByVal pnEstadoCheque As ChequeEstado, _
                                    ByVal psCodCtaAho As String, ByVal pdFecSis As Date) As Integer
Dim oMov As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim lnMovNro As Long

On Error GoTo GrabaExtornoChequeErr
Set oMov = New COMDMov.DCOMMov
GrabaExtornoCheque = 1
'oMov.inicio psFormatoFecha
oMov.BeginTrans
lbTrans = True
'oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable
'lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.ExtornaMovimiento psMovNro, pnMovNroAnt, psOpeCod, psMovDesc, True
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovOpeVarias lnMovNro, psNumCheque, Left(psOpeDesc, 12), pnImporte
Select Case psOpeCod
    Case gOpeChequesExtAnulación, gOpeChequesValorización
        oMov.ActualizaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, , , , , , , , , pnEstadoCheque
        oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pdFecSis, pnEstadoCheque, psMovNro
    Case gOpeChequesExtValorización
        oMov.ActualizaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, , , , , , , , , pnEstadoCheque
        oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pdFecSis, pnEstadoCheque, psMovNro
        'falta actualizar el saldo de la cuenta de ahorros que afecto el cheque
        If psCodCtaAho <> "" Then
        End If
End Select
'oMov.InsertaMovRef lnMovNro, pnMovNroAnt
GrabaExtornoCheque = 0
oMov.CommitTrans
lbTrans = False
Exit Function
GrabaExtornoChequeErr:
    If lbTrans Then
        lbTrans = False
        oMov.RollbackTrans
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub RegistroNotasAbonoCargo(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pnEstado As NotaCargoAbonoEstado, _
                                    ByVal pnMotivo As Integer, ByVal pnMonto As Currency, ByVal psMovNro As String, _
                                    Optional psObjetoCod As String = "", Optional psObjeto As String = "", Optional pnMoneda As Moneda = gMonedaNacional)
        
Dim oMov As COMDMov.DCOMMov
Dim lbTrans As Boolean
On Error GoTo RegistroNotasAbonoCargoErr
Set oMov = New COMDMov.DCOMMov
oMov.BeginTrans
lbTrans = True

oMov.InsertaNotaAbonoCargo pnTpoDoc, psNroDoc, pnEstado, pnMotivo, pnMonto, psObjetoCod, psObjeto, pnMoneda
oMov.InsertaNotaAbonoCargoEst pnTpoDoc, psNroDoc, pnEstado, psMovNro

oMov.CommitTrans
lbTrans = False
Exit Sub
RegistroNotasAbonoCargoErr:
    If lbTrans Then
        lbTrans = False
        oMov.RollbackTrans
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function GetMotivosNivel(ByVal pnTpoDoc As TpoDoc, Optional ByVal pnCodMotivo As MotivoNotaAbonoCargo = -1) As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
Dim lsFiltro As String
Set rs = New adodb.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
lsFiltro = ""
If pnCodMotivo <> -1 Then
    lsFiltro = " AND M.nMotivoCod='" & pnCodMotivo & "'"
End If
If oCon.AbreConexion = False Then Exit Function
sql = "Select M.nMotivoCod, M.cMotivoDesc, 1 as Nivel " _
    & " From MotivosNANC M JOIN DocMotivoNANC D ON D.NMOTIVOCOD = M.NMOTIVOCOD  " _
    & " WHERE D.nTpoDoc = " & pnTpoDoc & lsFiltro

Set rs = oCon.CargaRecordSet(sql)
Set GetMotivosNivel = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetNotasCargoAbonoEst(ByVal pnTpoDoc As TpoDoc, Optional ByVal pnEstado As NotaCargoAbonoEstado = -1, Optional ByVal pnMoneda As Moneda = gMonedaNacional) As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
Set rs = New adodb.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltroEst As String
Set oCon = New COMConecta.DCOMConecta
If pnEstado <> -1 Then
    lsFiltroEst = " and nEstado=" & pnEstado & " "
End If

If oCon.AbreConexion = False Then Exit Function
sql = "Select   Convert(char(20), cDocNro) + ' ' + Convert(varchar(30),cMotivoDesc)  as Documento        " _
    & " From    NotaAbonoCargo N JOIN MOTIVOSNANC  MN ON MN.nMotivoCod =N.nMotivoCod " _
    & " Where   nDocTpo =' " & pnTpoDoc & "' and nMoneda =" & pnMoneda & " " & lsFiltroEst

Set rs = oCon.CargaRecordSet(sql)
Set GetNotasCargoAbonoEst = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetDatosNotaAC(ByVal pnTpoDoc As TpoDoc, ByVal psNroNANC As String, Optional ByVal pnEstado As NotaCargoAbonoEstado = -1) As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
Set rs = New adodb.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltroEst As String
Set oCon = New COMConecta.DCOMConecta
If pnEstado <> -1 Then
    lsFiltroEst = " and nEstado=" & pnEstado & " "
End If

If oCon.AbreConexion = False Then Exit Function
sql = "select   nDocTpo, cDocNro, nEstado , nMotivoCod , cObjetoCodPadre, cObjetoCod, nMonto  " _
    & " From    NotaAbonoCargo " _
    & " Where   nDocTpo =' " & pnTpoDoc & "' and cDocNro ='" & psNroNANC & "' " & lsFiltroEst

Set rs = oCon.CargaRecordSet(sql)
Set GetDatosNotaAC = rs
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function GetMotivosObjNivel(ByVal pnMotivoCod As Long) As adodb.Recordset
Dim sql As String
Dim rs As adodb.Recordset
Set rs = New adodb.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function
sql = " Select  M.cObjetoCod, O.cObjetoDesc , 1 AS NIVEL " _
    & " From    MotivoNANCObj M JOIN OBJETO O ON O.COBJETOCOD = M.COBJETOCOD " _
    & " Where   nMotivocod =" & pnMotivoCod & ""

Set rs = oCon.CargaRecordSet(sql)
Set GetMotivosObjNivel = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetNroNotaCargoAbono(ByVal pnTpoDoc As TpoDoc) As String
Dim sql As String
Dim rs As adodb.Recordset
Set rs = New adodb.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function

sql = " Select Isnull(Max(cDocNro),0)  as MaxNota " _
    & " From NotaAbonoCargo " _
    & " Where  nDocTpo = '" & pnTpoDoc & "'"

Set rs = oCon.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetNroNotaCargoAbono = Format(rs!MaxNota + 1, "00000000")
Else
    GetNroNotaCargoAbono = Format(rs!MaxNota + 1, "00000000")
End If
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetDetalleObjetos(ByVal pnObjetoCod As TpoObjetos, Optional ByVal psFiltro As String = "", Optional pnNivel As MuestraIF = MuestraCuentas) As adodb.Recordset
Dim rs1 As New adodb.Recordset
Dim oRHAreas As COMDConstantes.DCOMActualizaDatosArea 'DActualizaDatosArea
Dim oContFunct As COMNContabilidad.NCOMContFunciones 'NContFunciones
Dim oCtaIf As COMNCajaGeneral.NCOMCajaCtaIF
Dim oEfect As COMDCajaGeneral.DCOMEfectivo
Dim lsFiltro As String

Set oCtaIf = New COMNCajaGeneral.NCOMCajaCtaIF
Set oContFunct = New COMNContabilidad.NCOMContFunciones
Set oRHAreas = New COMDConstantes.DCOMActualizaDatosArea
Set oEfect = New COMDCajaGeneral.DCOMEfectivo

Select Case Val(pnObjetoCod)
    Case ObjCMACAgencias
        Set rs1 = oRHAreas.GetAgencias(psFiltro)
    Case ObjCMACAgenciaArea
        Set rs1 = oRHAreas.GetAgenciasAreas(psFiltro)
    Case ObjCMACArea
        Set rs1 = oRHAreas.GetAreas(psFiltro)
    Case ObjEntidadesFinancieras
        Set rs1 = oCtaIf.GetCtasInstFinancieras(, , pnNivel, Val(psFiltro))
    Case ObjDescomEfectivo
        Set rs1 = oEfect.GetBilletajes(psFiltro)
    Case objPersona
        Set rs1 = Nothing
    Case Else
        Set rs1 = oContFunct.GetObjetos(psFiltro)
End Select
Set oCtaIf = Nothing
Set oContFunct = Nothing
Set oRHAreas = Nothing
Set oEfect = Nothing

Set GetDetalleObjetos = rs1
End Function

Public Function GetFiltroMotivoObj(ByVal pnMotivoCod As Long, ByVal psCodObjeto As String) As String
Dim sql As String
Dim rs As adodb.Recordset
Set rs = New adodb.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function
sql = " Select  M.cObjetoCod, M.cFiltro " _
    & " From    MotivoNANCObj M " _
    & " Where   M.nMotivocod =" & pnMotivoCod & " and cObjetoCod='" & psCodObjeto & "'"
GetFiltroMotivoObj = ""
Set rs = oCon.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetFiltroMotivoObj = rs!cFiltro
End If
rs.Close
Set rs = Nothing

oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetChequesNoDepositados(ByVal pnMoneda As Moneda, Optional ByVal psAreaCod As String = "", Optional psAgeCod As String = "") As adodb.Recordset
Dim sql As String
Dim rs As New adodb.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltro As String


Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function

lsFiltro = ""
If psAreaCod <> "" Or psAgeCod <> "" Then
    lsFiltro = "  AND DR.CAREACOD ='" & psAreaCod & "' AND DR.CAGECOD='" & psAgeCod & "' "
End If

sql = "Select   DISTINCT 0 as OK, P.cPersNombre AS Banco, DR.cNroDoc,DR.cIFCta,  " _
    & "         CONVERT(VARCHAR(10),DR.dValorizacion,103) AS FECHA, DR.nMonto, " _
    & "         ISNULL(C.cConsDescripcion,'') AS EstActual, ISNULL(DRC.cCtaCod,'') AS CuentaAhorros, " _
    & "         ISNULL(AG.cAgeDescripcion,A.cAreaDescripcion) AS AreaAgencia, DR.cAreaCod,  DR.cAgeCod, " _
    & "         CASE WHEN DRC.cCtaCod IS NOT NULL THEN O.COBJETOCOD ELSE OBJ.COBJETOCOD END AS OBJETO, M.NMOVNRO, DR.cPersCod, DR.cIFTpo " _
    & " From    DOCREC DR " _
    & "         JOIN MOVDOC MD ON  MD.nDocTpo= DR.nTpoDoc  AND DR.cNroDoc=MD.cDocNro " _
    & "         JOIN MOV    M  ON M.NMOVNRO = MD.NMOVNRO " _
    & "         LEFT JOIN ( SELECT  MC.nMovNro, MO.COBJETOCOD " _
    & "                     FROM    MOVCTA MC JOIN MOVOBJ MO ON MO.NMOVNRO=MC.NMOVNRO AND MC.NMOVITEM=MO.NMOVITEM " _
    & "                     WHERE   MO.COBJETOCOD  LIKE '60%') AS OBJ " _
    & "         ON OBJ.NMOVNRO =M.NMOVNRO " _
    & "         JOIN PERSONA P          ON P.cPersCod =DR.cPersCod " _
    & "         LEFT JOIN CONSTANTE C       ON C.nConsValor= DR.nEstado AND C.nConsCod =1007 " _
    & "         LEFT JOIN DocRecCapta DRC   ON DRC.nTpoDoc=DR.nTpoDoc AND DRC.cNrodoc=DR.cNrodoc " _
    & "                                         AND DRC.cPersCod=DR.cPersCod AND DRC.cIFTpo=DR.cIFTpo " _
    & "         JOIN AREAS A ON A.CAREACOD =DR.CAREACOD " _
    & "         LEFT JOIN AGENCIAS AG ON AG.CAGECOD =DR.CAGECOD " _
    & "         LEFT JOIN OBJETO O ON SUBSTRING(O.COBJETOCOD,5,3)=SUBSTRING(DRC.CCTACOD,6,3) " _
    & " WHERE   DR.cDepIF =" & gCGEstadosChqRecibido & " AND DR.nConfCaja=0 AND " _
    & "         DR.nEstado NOT IN (" & gChqEstAnulado & "," & gChqEstRechazado & "," & gsChqEstExtornado & ") " _
    & "         AND DR.nMoneda=" & pnMoneda & " AND M.NMOVFLAG=" & gMovFlagVigente & lsFiltro

Set rs = oCon.CargaRecordSet(sql)
Set GetChequesNoDepositados = rs
oCon.CierraConexion
Set oCon = Nothing
End Function


'EJVG20131216 ***
Public Function ListaIFis() As adodb.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim lsSql As String
    oCon.AbreConexion
    lsSql = "Exec stp_sel_ERS1262013_ListaIFisxCheque"
    Set ListaIFis = oCon.CargaRecordSet(lsSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function ListaCreditosxRecepcionCheque(ByVal psPersCod As String, ByVal pnMoneda As Moneda) As adodb.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim lsSql As String
    oCon.AbreConexion
    lsSql = "Exec stp_sel_ERS1262013_ListaCreditosxRecepcionCheque '" & psPersCod & "','" & pnMoneda & "'"
    Set ListaCreditosxRecepcionCheque = oCon.CargaRecordSet(lsSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function ExisteCheque(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psIFCta As String, ByVal psCheckDigit As String) As Boolean
    Dim rs As New adodb.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    Dim lsSql As String

    lsSql = "Exec stp_sel_ERS1262013_ExisteCheque " & pnTpoDoc & ",'" & psNroDoc & "','" & psPersCod & "','" & Format(psIFTpo, "00") & "','" & psIFCta & "','" & psCheckDigit & "'"
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(lsSql)
    If Not rs.EOF Then
        ExisteCheque = rs!bExiste
    End If
    oConect.CierraConexion
    Set rs = Nothing
    Set oConect = Nothing
End Function
Public Function RegistrarCheque(ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal psUser As String, _
                                ByVal psNroCheque As String, ByVal psNroCheque1 As String, ByVal psNroCheque2 As String, ByVal psNroCheque3 As String, _
                                ByVal psNroChequeCtaIF As String, ByVal psNroCheque4 As String, _
                                ByVal psTipoIF As String, ByVal psPersCodIF As String, _
                                ByVal psPersCodGirador As String, ByVal psPersCodContacto As String, _
                                ByVal pnImporte As Currency, ByVal pdFechaReg As Date, ByVal pdFechaVal As Date, _
                                ByVal pnMoneda As Moneda, ByVal pbChequeGerencia As Boolean, ByRef pMatOperaciones As Variant, _
                                ByVal pdDepFechaReg As Date, ByVal psDepIFTpo As String, ByVal psDepPersCod As String, ByVal psDepCtaIF As String, ByVal pnMovNroPend As Long, _
                                ByRef pMatMovNro() As String, ByVal psCtaContRegistroD As String, ByVal psCtaContRegistroH As String, ByVal psCtaContDepositoD As String, ByVal psCtaContDepositoH As String, _
                                Optional ByVal psNombreBenef As String = "", Optional ByVal pnDocTpoBenef As Integer = 0, Optional ByVal psDocNroBenef As String = "", Optional ByVal psNombreBcoDest As String = "", Optional ByVal pnMismoTit As Integer = 0, Optional ByVal pnEnvioCamara As Integer = 0) As Long
                                'psNombreBenef,pnDocTpoBenef,psDocNroBenef,psNombreBcoDest, pnMismoTit,pnEnvioCamara x PASI20160610 CCE
    Dim oMov As New COMDMov.DCOMMov
    Dim lbTrans As Boolean
    Dim lnMovNro As Long
    Dim lsMovNro As String
    Dim i As Integer
    Dim lsCtaIF As String
    Dim lnId As Long
    Dim lnIDRC As Long 'PASI20160610 CCE
    
    Dim lsChequeNro As String
    'Dim lnChequeTpoOpe As Integer
    Dim lnChequeTpoOpe As TipoOperacionCheque
    Dim lnChequeMonto As Currency
    Dim lsChequeEstado As String
    Dim lsChequeDet As String
    Dim lsChequeGlosa As String
    Dim bDepositoEnRegistro As Boolean
    
    On Error GoTo ErrRegistrarCheque
    oMov.BeginTrans
    lbTrans = True
    lsCtaIF = psNroChequeCtaIF
    bDepositoEnRegistro = IIf(DateDiff("d", pdDepFechaReg, CDate("01/01/1900")) <> 0, True, False)
        
    lnId = oMov.InsertaChequeCabecera(TpoDocCheque, psNroCheque, psPersCodIF, psTipoIF, lsCtaIF, pnImporte, pdFechaReg, pdFechaVal, pnMoneda, psAreaCod, psAgeCod, psUser, psPersCodGirador, psPersCodContacto, 0, pbChequeGerencia, psNroCheque1, psNroCheque2, psNroCheque3, psNroCheque4, bDepositoEnRegistro)
     
    If lnId = 0 Then
        oMov.RollbackTrans
        Set oMov = Nothing
        Exit Function
    End If
    
    If Not pnEnvioCamara = 0 Then oMov.CCE_InsertaCheque lnId, psNombreBenef, pnDocTpoBenef, psDocNroBenef, psNombreBcoDest, pnMismoTit 'PASI20160610 CCE
    
    ReDim pMatMovNro(0)
    For i = 1 To UBound(pMatOperaciones, 2)
        lsChequeNro = psNroCheque & "-" & CStr(i) 'pMatOperaciones(1, i)
        lnChequeTpoOpe = CInt(Trim(Right(pMatOperaciones(2, i), 2)))
        lnChequeMonto = pMatOperaciones(3, i)
        lsChequeEstado = pMatOperaciones(4, i)
        lsChequeDet = pMatOperaciones(5, i)
        lsChequeGlosa = pMatOperaciones(6, i)

        Sleep 1000
        lsMovNro = oMov.GeneraMovNro(pdFechaReg, psAgeCod, psUser)
        oMov.InsertaMov lsMovNro, "900031", lsChequeGlosa, gMovEstContabMovContable, gMovFlagVigente
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        oMov.InsertaMovDoc lnMovNro, TpoDocCheque, lsChequeNro, Format(pdFechaReg, "mm/dd/yyyy")
        oMov.InsertaMovDocRec lnMovNro, TpoDocCheque, lsChequeNro, psPersCodIF, psTipoIF, lsCtaIF
        oMov.InsertaCheque TpoDocCheque, lsChequeNro, psPersCodIF, psTipoIF, 0, lsCtaIF, lnChequeMonto, pdFechaVal, _
                                pdFechaVal, gCGEstadosChqRecibido, 0, pnMoneda, gChqEstRegistrado, psAreaCod, psAgeCod, 0, psPersCodGirador, _
                                psNroCheque & "-" & psNroCheque1 & "-" & psNroCheque2 & "-" & psNroCheque3 & "-" & lsCtaIF & "-" & psNroCheque4, lnMovNro, pbChequeGerencia
        oMov.InsertaChequeEstado TpoDocCheque, lsChequeNro, psPersCodIF, psTipoIF, pdFechaReg, gChqEstRegistrado, lsMovNro, lsCtaIF
        If Not bDepositoEnRegistro Then
            oMov.InsertaMovCta lnMovNro, 1, psCtaContRegistroD, lnChequeMonto
            oMov.InsertaMovCta lnMovNro, 2, psCtaContRegistroH, lnChequeMonto * -1
        Else
            oMov.InsertaMovCta lnMovNro, 1, psCtaContDepositoD, lnChequeMonto
            If pnMovNroPend = 0 Then 'Si es en el día si muestra en libro banco
                oMov.InsertaMovObj lnMovNro, 1, 1, ObjEntidadesFinancieras
                oMov.InsertaMovObjIF lnMovNro, 1, 1, psDepPersCod, psDepIFTpo, psDepCtaIF
            End If
            oMov.InsertaMovCta lnMovNro, 2, psCtaContRegistroH, lnChequeMonto * -1
        End If
        
'        Select Case lnChequeTpoOpe
'            Case 1 'Apertura
'                oMov.ActualizaChequexCab TpoDocCheque, lsChequeNro, psPersCodIF, psTipoIF, lsCtaIF, lnChequeTpoOpe, lsChequeDet
'            Case 2, 3, 4, 6, 10 'Deposito, Aumento Capital y Pago Crédito'WIOR 20160412 agrego codigo 10 liquidacion cred seg. desg
'                oMov.ActualizaChequexCab TpoDocCheque, lsChequeNro, psPersCodIF, psTipoIF, lsCtaIF, lnChequeTpoOpe, , lsChequeDet
'            Case 5, 7, 8 'Nro Clientes
'                oMov.ActualizaChequexCab TpoDocCheque, lsChequeNro, psPersCodIF, psTipoIF, lsCtaIF, lnChequeTpoOpe, , , CInt(lsChequeDet)
'        End Select
        Select Case lnChequeTpoOpe
            Case DPF_Apertura, AHO_Apertura, CTS_Apertura 'Apertura
                oMov.ActualizaChequexCab TpoDocCheque, lsChequeNro, psPersCodIF, psTipoIF, lsCtaIF, lnChequeTpoOpe, lsChequeDet
            Case DPF_AumentoCapital, AHO_Deposito, CTS_Deposito, CRED_Pago, CRED_LiqCreditosSegDes 'Deposito, Aumento Capital y Pago Crédito
                oMov.ActualizaChequexCab TpoDocCheque, lsChequeNro, psPersCodIF, psTipoIF, lsCtaIF, lnChequeTpoOpe, , lsChequeDet
            Case CTS_DepositoLote, CRED_PagoLote, AHO_DepositoLote, AHO_AperturaLote, AHO_DepositoHaberesLote, DPF_AperturaLote 'Nro Clientes
                oMov.ActualizaChequexCab TpoDocCheque, lsChequeNro, psPersCodIF, psTipoIF, lsCtaIF, lnChequeTpoOpe, , , CInt(lsChequeDet)
        End Select
        oMov.InsertaChequeRel lnId, TpoDocCheque, lsChequeNro, psPersCodIF, psTipoIF, lsCtaIF
        If pnMovNroPend > 0 Then
            oMov.InsertaMovRef lnMovNro, pnMovNroPend
        End If
        If pnMoneda = gMonedaExtranjera Then
            oMov.GeneraMovME lnMovNro, lsMovNro
        End If
        oMov.ActualizaSaldoMovimiento lsMovNro, "+"
        If pnMovNroPend > 0 Then 'Revertimos asiento pendiente x Deposito antes de la fecha de recepción
            oMov.RevierteAsientoPendientexRecepcionCheque lnMovNro, pnMovNroPend, lnChequeMonto
        End If
        ReDim Preserve pMatMovNro(0 To i)
        pMatMovNro(i) = lsMovNro
    Next
    If bDepositoEnRegistro Then
        oMov.RegistraDepositoCheque lnId, pdDepFechaReg, psDepIFTpo, psDepPersCod, psDepCtaIF
    End If

    oMov.CommitTrans
    lbTrans = False
    RegistrarCheque = lnId
    Set oMov = Nothing
    Exit Function
ErrRegistrarCheque:
    RegistrarCheque = 0
    If lbTrans Then
        oMov.RollbackTrans
        Set oMov = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'Public Function EditarCheque(ByVal pnId As Long, ByVal psPersCodGirador As String, ByVal psPersCodContacto As String, ByRef pMatOperaciones As Variant) As Long /**Comentado PASI20161212 CCE**/
Public Function EditarCheque(ByVal pnId As Long, ByVal psPersCodGirador As String, ByVal psPersCodContacto As String, ByRef pMatOperaciones As Variant, Optional ByVal pbEsChequexAjuste As Boolean = False, Optional ByVal pnMonto As Currency = 0, Optional ByVal pnMontoAnterior As Currency = 0, Optional ByVal pdFechaReg As Date = "01/01/1900", Optional ByVal psAgeCod As String = "", Optional ByVal psCodUser As String = "", Optional ByVal psNroCheque As String = "", Optional pnMoneda As Integer = 0, Optional ByVal psCtaContD As String = "", Optional ByVal psCtaContH As String = "") As Long 'PASI20161212 CCE
    Dim oMov As New COMDMov.DCOMMov
    Dim lbTrans As Boolean
    Dim lsMovNro As String 'PASI20161212 CCE
    Dim lnMovNro As Long 'PASI20161212 CCE
    Dim i As Integer
    Dim lsChequeNro As String
    Dim lsChequeTpoOpe As Integer
    Dim lsChequeMonto As Currency
    Dim lsChequeEstado As String
    Dim lsChequeDet As String
    Dim lsChequeGlosa As String
    Dim rsCheque As New adodb.Recordset
    Dim bActualizaDetalle As Boolean
    
    On Error GoTo ErrRegistrarCheque
    
    oMov.BeginTrans
    lbTrans = True

    Set rsCheque = ChequexID(pnId)
    'oMov.ActualizaChequeCabecera pnId, psPersCodGirador, psPersCodContacto /**Comentado PASI20161212 CCE**/
    oMov.ActualizaChequeCabecera pnId, psPersCodGirador, psPersCodContacto, pnMonto 'PASI20161212 CCE
        
    For i = 1 To UBound(pMatOperaciones, 2)
        lsChequeNro = pMatOperaciones(1, i)
        lsChequeTpoOpe = CInt(Trim(Right(pMatOperaciones(2, i), 2)))
        lsChequeMonto = pMatOperaciones(3, i)
        lsChequeEstado = pMatOperaciones(4, i)
        lsChequeDet = pMatOperaciones(5, i)
        lsChequeGlosa = pMatOperaciones(6, i)

        If lsChequeEstado = "PENDIENTE" Then
            bActualizaDetalle = True
            Select Case lsChequeTpoOpe
                Case 1, 11, 14  'Apertura
                    'oMov.ActualizaChequexCab TpoDocCheque, lsChequeNro, rsCheque!cPersCod, rsCheque!cIFTpo, rsCheque!cIFCta, lsChequeTpoOpe, lsChequeDet /**Comentado PASI20161212 CCE**/
                    oMov.ActualizaChequexCab TpoDocCheque, lsChequeNro, rsCheque!cPersCod, rsCheque!cIFTpo, rsCheque!cIFCta, lsChequeTpoOpe, lsChequeDet, , , pbEsChequexAjuste, lsChequeMonto 'PASI20161212 CCE
                Case 2, 3, 4, 6 'Deposito, Aumento Capital y Pago Crédito
                    'oMov.ActualizaChequexCab TpoDocCheque, lsChequeNro, rsCheque!cPersCod, rsCheque!cIFTpo, rsCheque!cIFCta, lsChequeTpoOpe, , lsChequeDet /**Comentado PASI20161212 CCE**/
                    oMov.ActualizaChequexCab TpoDocCheque, lsChequeNro, rsCheque!cPersCod, rsCheque!cIFTpo, rsCheque!cIFCta, lsChequeTpoOpe, , lsChequeDet, , pbEsChequexAjuste, lsChequeMonto 'PASI20161212 CCE
                Case 5, 7, 8, 12, 13, 15 'Nro Clientes
                    'oMov.ActualizaChequexCab TpoDocCheque, lsChequeNro, rsCheque!cPersCod, rsCheque!cIFTpo, rsCheque!cIFCta, lsChequeTpoOpe, , , CInt(lsChequeDet) /**Comentado PASI20161212 CCE**/
                    oMov.ActualizaChequexCab TpoDocCheque, lsChequeNro, rsCheque!cPersCod, rsCheque!cIFTpo, rsCheque!cIFCta, lsChequeTpoOpe, , , CInt(lsChequeDet), pbEsChequexAjuste, lsChequeMonto 'PASI20161212 CCE
                Case Else
                    bActualizaDetalle = False
            End Select
            If bActualizaDetalle Then
                oMov.ActualizaDatosMovimientoCheque TpoDocCheque, lsChequeNro, rsCheque!cPersCod, rsCheque!cIFTpo, rsCheque!cIFCta, lsChequeGlosa
            End If
        End If
    Next
    'PASI20161212 CCE*********************************
    If pbEsChequexAjuste Then
        lsMovNro = oMov.GeneraMovNro(pdFechaReg, psAgeCod, psCodUser)
        oMov.InsertaMov lsMovNro, "900040", "Ajuste de Cheque  N° " & psNroCheque, gMovEstContabMovContable, gMovFlagVigente
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        oMov.InsertaMovCta lnMovNro, 1, psCtaContD, (pnMonto - pnMontoAnterior)
        oMov.InsertaMovCta lnMovNro, 2, psCtaContH, (pnMontoAnterior - pnMonto)
        If pnMoneda = gMonedaExtranjera Then
            oMov.GeneraMovME lnMovNro, lsMovNro
        End If
        oMov.CCE_InsertaChequeAjuste pnId, pnMontoAnterior, pnMonto, lnMovNro
        oMov.ActualizaSaldoMovimiento lsMovNro, "+"
    End If
    'PASI END*****************************************

    oMov.CommitTrans
    lbTrans = False
    EditarCheque = pnId
    
    Set rsCheque = Nothing
    Set oMov = Nothing
    Exit Function
ErrRegistrarCheque:
    EditarCheque = 0
    If lbTrans Then
        oMov.RollbackTrans
        Set oMov = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function ChequexImpresion(ByVal pnId As Long) As adodb.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ChequexImpresion " & pnId
    oConect.AbreConexion
    Set ChequexImpresion = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function ListaChequexMantenimiento(ByVal psAgeCod As String) As adodb.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ListaChequexMantenimiento '" & psAgeCod & "'"
    oConect.AbreConexion
    Set ListaChequexMantenimiento = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function ChequexID(ByVal pnId As Long) As adodb.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ChequexId " & pnId
    oConect.AbreConexion
    Set ChequexID = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function ChequeDetxID(ByVal pnId As Long) As adodb.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ChequeDetxId " & pnId
    oConect.AbreConexion
    Set ChequeDetxID = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function ListaChequexExtorno(ByVal psNroDoc As String, ByVal psAgeCod As String, ByVal pdFecha As Date) As adodb.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ListaChequexExtorno '" & psNroDoc & "','" & psAgeCod & "','" & Format(pdFecha, "yyyymmdd") & "'"
    oConect.AbreConexion
    Set ListaChequexExtorno = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
'*****cti3
Public Function ExtornarCheque(ByVal pnId As Long, ByVal psMovNro As String, _
ByVal psGlosa As String, Optional ByVal psMotExtorno As Variant) As Boolean
    Dim oMov As New COMDMov.DCOMMov
    Dim lbTrans As Boolean
    Dim lnMovNro As Long

    On Error GoTo ErrExtornarCheque
    oMov.BeginTrans
    lbTrans = True
    oMov.InsertaMov psMovNro, gChqOpeExtRegistro, Left(psGlosa, 300), gMovEstContabNoContable, gMovFlagDeExtorno
    lnMovNro = oMov.GetnMovNro(psMovNro)
    oMov.InsertaDetExtorno lnMovNro, gChqOpeExtRegistro, "", psMotExtorno 'CTI3  x
    
    oMov.ExtornarCheque pnId, lnMovNro
    oMov.CommitTrans
    lbTrans = False
    ExtornarCheque = True
    Set oMov = Nothing
    Exit Function
ErrExtornarCheque:
    ExtornarCheque = False
    If lbTrans Then
        oMov.RollbackTrans
        Set oMov = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function ListaChequexDeposito(ByVal psAgeCod As String, ByVal pnMoneda As Integer) As adodb.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ListaChequexDeposito '" & psAgeCod & "'," & pnMoneda
    oConect.AbreConexion
    Set ListaChequexDeposito = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function RegistrarDepositoCheque(ByRef pMatDatos() As Long, ByVal pnMoneda As Moneda, ByVal psIFTpo As String, ByVal psPersCod As String, ByVal psCtaIfCod As String, _
                                        ByVal pdFecha As Date, ByVal psAgeCod As String, ByVal psUserCod As String, _
                                        ByVal psCtaContDepositoD As String, ByVal psCtaContDepositoH As String, ByRef psMovNroAsiento As String) As Boolean
    Dim oMov As New COMDMov.DCOMMov
    Dim oRSDet As New adodb.Recordset
    Dim lbTrans As Boolean
    Dim lsMovNro As String
    Dim lnMovNro As Long, lnMovNroRef As Long, lnId As Long
    Dim i As Long
    
    On Error GoTo ErrRegistrarDepositoCheque
    oMov.BeginTrans
    lbTrans = True
    psMovNroAsiento = ""
    
    For i = 1 To UBound(pMatDatos)
        lnId = pMatDatos(i)
        Set oRSDet = ChequeDetxID(lnId)
        If Not oRSDet.EOF Then
            Do While Not oRSDet.EOF
                lnMovNroRef = oRSDet!nMovNro
                Sleep 1000
                lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, psUserCod)
                oMov.InsertaMov lsMovNro, gChqOpeDeposito, "Depósito Cheque Nro. " & oRSDet!cNroDoc, gMovEstContabMovContable, gMovFlagVigente
                lnMovNro = oMov.GetnMovNro(lsMovNro)
                oMov.InsertaMovCta lnMovNro, 1, psCtaContDepositoD, Abs(oRSDet!nMonto)
                oMov.InsertaMovObj lnMovNro, 1, 1, ObjEntidadesFinancieras
                oMov.InsertaMovObjIF lnMovNro, 1, 1, psPersCod, psIFTpo, psCtaIfCod
                oMov.InsertaMovCta lnMovNro, 2, psCtaContDepositoH, Abs(oRSDet!nMonto) * -1
                oMov.InsertaMovRef lnMovNro, lnMovNroRef
                oMov.InsertaMovDoc lnMovNro, oRSDet!nTpoDoc, oRSDet!cNroDoc, Format(pdFecha, "yyyymmdd hh:mm:ss")
                oMov.InsertaMovDocRec lnMovNro, oRSDet!nTpoDoc, oRSDet!cNroDoc, oRSDet!cPersCod, oRSDet!cIFTpo, oRSDet!cIFCta
                oMov.ActualizaNroMovDeposito lnMovNroRef, lnMovNro
                
                If pnMoneda = gMonedaExtranjera Then
                    oMov.GeneraMovME lnMovNro, lsMovNro
                End If
                oMov.ActualizaSaldoMovimiento lsMovNro, "+"
                psMovNroAsiento = psMovNroAsiento & lsMovNro & ","
                oRSDet.MoveNext
            Loop
            oMov.RegistraDepositoCheque lnId, pdFecha, psIFTpo, psPersCod, psCtaIfCod
        End If
    Next
    If Len(psMovNroAsiento) > 0 Then
        psMovNroAsiento = Left(psMovNroAsiento, Len(psMovNroAsiento) - 1)
    End If

    oMov.CommitTrans
    lbTrans = False
    RegistrarDepositoCheque = True
    Set oMov = Nothing
    Exit Function
ErrRegistrarDepositoCheque:
    RegistrarDepositoCheque = False
    If lbTrans Then
        oMov.RollbackTrans
        Set oMov = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function ListaIFiChequexBusqueda(ByVal pnMoneda As Moneda, ByVal psAgeCod As String, ByVal pnTpoOperacion As TipoOperacionCheque, ByVal psDetalle As String) As adodb.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ListaIFiChequexBusqueda " & pnMoneda & ",'" & psAgeCod & "'," & pnTpoOperacion & ",'" & psDetalle & "'"
    oConect.AbreConexion
    Set ListaIFiChequexBusqueda = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function ListaChequexBusqueda(ByVal pnMoneda As Moneda, ByVal psAgeCod As String, ByVal pnTpoOperacion As TipoOperacionCheque, ByVal psPersCod As String, ByVal psDetalle As String) As adodb.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ListaChequexBusqueda " & pnMoneda & ",'" & psAgeCod & "'," & pnTpoOperacion & ",'" & psPersCod & "','" & psDetalle & "'"
    oConect.AbreConexion
    Set ListaChequexBusqueda = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function ChequexBusqueda(ByVal pnTpoDoc As Integer, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psIFCta As String) As adodb.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ChequexBusqueda " & pnTpoDoc & ",'" & psNroDoc & "','" & psPersCod & "','" & Format(psIFTpo, "00") & "','" & psIFCta & "'"
    oConect.AbreConexion
    Set ChequexBusqueda = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function ListaChequexDevSobranteOperacion(ByVal psPersCodTitCred As String) As adodb.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ListaChequexDevSobranteOperacion '" & psPersCodTitCred & "'"
    oConect.AbreConexion
    Set ListaChequexDevSobranteOperacion = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function ListaIFiChequexBusquedaCredConvenio(ByVal pnMoneda As Moneda, ByVal psPersCodGirador As String) As adodb.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ListaIFiChequexBusquedaCredConvenio " & pnMoneda & ",'" & psPersCodGirador & "'"
    oConect.AbreConexion
    Set ListaIFiChequexBusquedaCredConvenio = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function ListaChequexBusquedaCredConvenio(ByVal pnMoneda As Moneda, ByVal psPersCodGirador As String) As adodb.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ListaChequexBusquedaCredConvenio " & pnMoneda & ",'" & psPersCodGirador & "'"
    oConect.AbreConexion
    Set ListaChequexBusquedaCredConvenio = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function ListaChequexOperacionesPendiente(ByVal psAgeCod As String) As adodb.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ListaChequexOperacionesPendiente '" & psAgeCod & "'"
    oConect.AbreConexion
    Set ListaChequexOperacionesPendiente = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function ListaChequeDetxOperacionesPendiente(ByVal pnEstructuraNueva As Boolean, ByVal pnId As Long, ByVal pnTpoDoc As Integer, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psIFCta As String) As adodb.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ListaChequeDetxOperacionesPendiente " & IIf(pnEstructuraNueva, 1, 0) & "," & pnId & "," & pnTpoDoc & ",'" & psNroDoc & "','" & psPersCod & "','" & psIFTpo & "','" & psIFCta & "'"
    oConect.AbreConexion
    Set ListaChequeDetxOperacionesPendiente = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function ListaChequexExtornoDeposito(ByVal psNroDoc As String, ByVal psAgeCod As String, ByVal pdFecha As Date) As adodb.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ListaChequexExtornoDeposito '" & psNroDoc & "','" & psAgeCod & "','" & Format(pdFecha, "yyyymmdd") & "'"
    oConect.AbreConexion
    Set ListaChequexExtornoDeposito = oConect.CargaRecordSet(lsSql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function ExtornarDepositoCheque(ByVal pnId As Long, ByVal psMovNro As String, _
                                ByVal psGlosa As String, Optional ByVal psMotExtorno As Variant) As Boolean
    Dim oMov As New COMDMov.DCOMMov
    Dim lbTrans As Boolean
    Dim lnMovNro As Long
    Dim rsDet As New adodb.Recordset

    On Error GoTo ErrExtornarCheque
    oMov.BeginTrans
    lbTrans = True
    
    Set rsDet = ChequeDetxID(pnId)
    oMov.InsertaMov psMovNro, gChqOpeExtDeposito, Left(psGlosa, 300), gMovEstContabNoContable, gMovFlagDeExtorno
    lnMovNro = oMov.GetnMovNro(psMovNro)
    oMov.InsertaDetExtorno lnMovNro, gChqOpeExtRegistro, "", psMotExtorno 'CTI3 : Acta 186-2018
        
    Do While Not rsDet.EOF
        oMov.ActualizaMov rsDet!nMovNroDep, , , gMovFlagExtornado
        oMov.InsertaMovRef lnMovNro, rsDet!nMovNroDep
        rsDet.MoveNext
    Loop
    oMov.CommitTrans
    lbTrans = False
    ExtornarDepositoCheque = True
    Set oMov = Nothing
    Exit Function
ErrExtornarCheque:
    ExtornarDepositoCheque = False
    If lbTrans Then
        oMov.RollbackTrans
        Set oMov = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'END EJVG *******

'PASI20140618
Public Function ObtieneExisteDocCheque(ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psIFCta As String) As Boolean
    Dim oConect As New COMConecta.DCOMConecta
    Dim rs As adodb.Recordset
    Set rs = New adodb.Recordset
    Dim lsSql As String
    
    lsSql = "Exec stp_sel_ObtieneExisteDocCheque '" & psNroDoc & "','" & psPersCod & "','" & psIFTpo & "','" & psIFCta & "'"
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(lsSql)
    ObtieneExisteDocCheque = IIf(rs!nvalor = 0, False, True)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
'end PASI

