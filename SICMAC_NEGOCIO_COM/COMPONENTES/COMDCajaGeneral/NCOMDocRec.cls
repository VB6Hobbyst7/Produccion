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
            Optional pnProducto As Producto = 0) As Integer
            
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
oMov.InsertaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pnPlaza, psCtaCheque, pnImporte, pdFechaVal, _
                        pdFechaVal, pnEstadoCG, pnConfirmaCaja, pnMonedaCheque, gChqEstRegistrado, psAreaCod, psAgeCod, pnProducto
                        
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
                                      ByVal psCtaHaber As String, ByVal rsObjHaber As ADODB.Recordset, _
                                      ByVal psNumCheque As String, ByVal psPersCodIF As String, ByVal psTipoIF As String, ByVal pnPlaza As ChequePlaza, _
                                      ByVal psCtaCheque As String, ByVal pnImporte As Currency, ByVal pdFechaReg As Date, _
                                      ByVal pdFechaVal As Date, ByVal psFormatoFecha As String, ByVal pnMonedaCheque As Moneda, _
                                      Optional ByVal pnEstadoCheque As ChequeEstado = gChqEstEnValorizacion, _
                                      Optional ByVal pnEstadoCG As CGEstadosChq = gCGEstadosChqRecibido, _
                                      Optional ByVal pnConfirmaCaja As CGEstadoConfCheque = ChqCGSinConfirmacion, _
                                      Optional ByVal psAreaCodChq As String, Optional ByVal psAgeCodChq As String) As Integer

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
                        pdFechaVal, pnEstadoCG, pnConfirmaCaja, pnMonedaCheque, , psAreaCodChq, psAgeCodChq

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
Public Function GetDatosCheques(ByVal psNumChq As String, ByVal psPersCodIF As String) As ADODB.Recordset
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta

Set oConect = New COMConecta.DCOMConecta
Set Rs = New ADODB.Recordset

If oConect.AbreConexion = False Then Exit Function

Sql = "SELECT   DR.CPERSCOD , LEFT(P.CPERSNOMBRE,40) AS IFNOMBRE, " _
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

Set Rs = oConect.CargaRecordSet(Sql)
Set GetDatosCheques = Rs
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
sSql = "Select cPersNombre, cNroDoc, nMonto = nMonto - nMontoUsado, cPersCod, dValorizaRef From ( " _
    & "SELECT P.cPersNombre, D.cNroDoc, D.nMonto, P.cPersCod, dValorizaRef, SUM(ISNULL(DC.nMonto,0)) nMontoUsado " _
    & "FROM (Select DC.nMonto, DC.nTpoDoc, DC.cNroDoc, DC.cPersCod, DC.cIFTpo From Mov M JOIN DocRecCapta DC " _
    & "ON M.nMovNro = DC.nMovNro Where M.nMovFlag = 0) DC RIGHT JOIN DocRec D INNER JOIN DocRecEst E INNER JOIN Persona P ON E.cPersCod = P.cPersCod " _
    & "ON D.nTpoDoc = E.nTpoDoc AND D.cNroDoc = E.cNroDoc AND D.cPersCod = E.cPersCod AND D.cIFTpo = E.cIFTpo " _
    & "ON DC.nTpoDoc = D.nTpoDoc AND DC.cNroDoc = D.cNroDoc AND DC.cPersCod = D.cPersCod AND DC.cIFTpo = D.cIFTpo " _
    & "WHERE D.nEstado IN (" & gChqEstRegistrado & "," & gsChqEstExtornado & ") AND E.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' " _
    & "AND E.cMovNro IN (Select MAX(cMovNro) FROM DocRecEst E1 " _
    & "WHERE E1.nTpoDoc = E.nTpoDoc AND E1.cNroDoc = E.cNroDoc AND E1.cPersCod = E.cPersCod AND E1.cIFTpo = E.cIFTpo) " _
    & "And D.nMoneda = " & nMoneda & " Group by P.cPersNombre, D.cNroDoc, D.nMonto, P.cPersCod, dValorizaRef " _
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

sSql = " SELECT DISTINCT cPersNombre, cNroDoc, nMonto , D.cPersCod, dValorizaRef FROM DOCREC D  "
sSql = sSql & "  JOIN PERSONA P ON P.CPERSCOD=D.CPERSCOD "
sSql = sSql & "  JOIN MOVDOC MD ON MD.CDOCNRO=D.CNRODOC "
sSql = sSql & "  LEFT JOIN (SELECT * FROM MOV WHERE NMOVFLAG=0 AND COPECOD='200252' ) M ON M.NMOVNRO=MD.NMOVNRO "
sSql = sSql & "  WHERE D.NESTADO IN (1,2) AND D.NPRODUCTO NOT IN (233,234,232) AND M.NMOVNRO IS NULL  and D.NMONEDA=" & nMoneda


Set clsConecta = New COMConecta.DCOMConecta
If clsConecta.AbreConexion = False Then Exit Function
Set GetChequesCreditos = clsConecta.CargaRecordSet(sSql)
clsConecta.CierraConexion
Set clsConecta = Nothing

End Function


Public Function GetCheques(ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psAgeCod As String, Optional psNumCheque As String) As ADODB.Recordset
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltro As String


Set oCon = New COMConecta.DCOMConecta
lsFiltro = ""
If Trim(psNumCheque) <> "" Then
    lsFiltro = " AND DR.cNroDoc LIKE '%" & psNumCheque & "%'"
End If

If oCon.AbreConexion = False Then Exit Function

Sql = " Select  DISTINCT DR.cNroDoc, P.cPersNombre, ISNULL(DRC.cCtaCod,'') AS cCtaCod, DR.nMonto," _
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

Set Rs = oCon.CargaRecordSet(Sql)
Set GetCheques = Rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetPersCuentaAho(ByVal psCodCta As String) As ADODB.Recordset
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltro As String


Set oCon = New COMConecta.DCOMConecta
'lsFiltro = ""
'If Trim(psNumCheque) <> "" Then
'    lsFiltro = " AND DR.cNroDoc LIKE '%" & psNumCheque & "%'"
'End If

If oCon.AbreConexion = False Then Exit Function

Sql = "     Select  P.cPersNombre as Nombre, T1.cConsDescripcion  as Relacion," _
        & "         p.nPersPersoneria  as Personeria " _
        & " FROM    Persona P " _
        & "         JOIN ProductoPersona PP ON P.cPersCod = PP.cPersCod " _
        & "         JOIN Constante T1 ON PP.NPrdPersRelac = T1.nConsValor " _
        & " WHERE   PP.cCtaCod = '" & psCodCta & "' AND T1.nCONSCOD = " & gCaptacRelacPersona & " " _
        & " ORDER BY Nombre "

Set Rs = oCon.CargaRecordSet(Sql)
Set GetPersCuentaAho = Rs
oCon.CierraConexion
Set oCon = Nothing
End Function


Public Function GetEstadosCheques(ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF) As ADODB.Recordset
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function

Sql = "      Select  CONVERT(VARCHAR(12),Convert(datetime, Substring(DR.cMovNro,1,8)),103) as Fecha ," _
        & "          ISNULL(C.cConsDescripcion,'') AS EstActual " _
        & "  From    DocRecEst DR " _
        & "          LEFT JOIN CONSTANTE C        ON  C.nConsValor= DR.nEstado AND C.nConsCod =1007 " _
        & "  Where  DR.nTpoDoc = " & TpoDocCheque & " And DR.cNroDoc = '" & psNroDoc & "' " _
        & "         And DR.cPersCod ='" & psPersCod & "' And DR.cIFTpo='" & Format(psIFTpo, "00") & "' " _
        & " ORDER BY DR.cMovNro, DR.nEstado ASC  "

Set Rs = oCon.CargaRecordSet(Sql)
Set GetEstadosCheques = Rs
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
Public Function GetOpeCheques(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psAgeCod As String, Optional psNumCheque As String) As ADODB.Recordset
Dim Sql As String
Dim Rs As New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltro As String


Set oCon = New COMConecta.DCOMConecta
lsFiltro = ""
If Trim(psNumCheque) <> "" Then
    lsFiltro = " AND DR.cNroDoc LIKE '%" & psNumCheque & "%'"
End If

If oCon.AbreConexion = False Then Exit Function

Sql = " Select  CONVERT(VARCHAR(12), CONVERT(DATETIME,SUBSTRING(M.cMovNro,1,8)),103) AS FECHA, " _
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

Set Rs = oCon.CargaRecordSet(Sql)
Set GetOpeCheques = Rs
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
Public Function GetMotivosNivel(ByVal pnTpoDoc As TpoDoc, Optional ByVal pnCodMotivo As MotivoNotaAbonoCargo = -1) As ADODB.Recordset
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim lsFiltro As String
Set Rs = New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
lsFiltro = ""
If pnCodMotivo <> -1 Then
    lsFiltro = " AND M.nMotivoCod='" & pnCodMotivo & "'"
End If
If oCon.AbreConexion = False Then Exit Function
Sql = "Select M.nMotivoCod, M.cMotivoDesc, 1 as Nivel " _
    & " From MotivosNANC M JOIN DocMotivoNANC D ON D.NMOTIVOCOD = M.NMOTIVOCOD  " _
    & " WHERE D.nTpoDoc = " & pnTpoDoc & lsFiltro

Set Rs = oCon.CargaRecordSet(Sql)
Set GetMotivosNivel = Rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetNotasCargoAbonoEst(ByVal pnTpoDoc As TpoDoc, Optional ByVal pnEstado As NotaCargoAbonoEstado = -1, Optional ByVal pnMoneda As Moneda = gMonedaNacional) As ADODB.Recordset
Dim Sql As String
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltroEst As String
Set oCon = New COMConecta.DCOMConecta
If pnEstado <> -1 Then
    lsFiltroEst = " and nEstado=" & pnEstado & " "
End If

If oCon.AbreConexion = False Then Exit Function
Sql = "Select   Convert(char(20), cDocNro) + ' ' + Convert(varchar(30),cMotivoDesc)  as Documento        " _
    & " From    NotaAbonoCargo N JOIN MOTIVOSNANC  MN ON MN.nMotivoCod =N.nMotivoCod " _
    & " Where   nDocTpo =' " & pnTpoDoc & "' and nMoneda =" & pnMoneda & " " & lsFiltroEst

Set Rs = oCon.CargaRecordSet(Sql)
Set GetNotasCargoAbonoEst = Rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetDatosNotaAC(ByVal pnTpoDoc As TpoDoc, ByVal psNroNANC As String, Optional ByVal pnEstado As NotaCargoAbonoEstado = -1) As ADODB.Recordset
Dim Sql As String
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltroEst As String
Set oCon = New COMConecta.DCOMConecta
If pnEstado <> -1 Then
    lsFiltroEst = " and nEstado=" & pnEstado & " "
End If

If oCon.AbreConexion = False Then Exit Function
Sql = "select   nDocTpo, cDocNro, nEstado , nMotivoCod , cObjetoCodPadre, cObjetoCod, nMonto  " _
    & " From    NotaAbonoCargo " _
    & " Where   nDocTpo =' " & pnTpoDoc & "' and cDocNro ='" & psNroNANC & "' " & lsFiltroEst

Set Rs = oCon.CargaRecordSet(Sql)
Set GetDatosNotaAC = Rs
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function GetMotivosObjNivel(ByVal pnMotivoCod As Long) As ADODB.Recordset
Dim Sql As String
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function
Sql = " Select  M.cObjetoCod, O.cObjetoDesc , 1 AS NIVEL " _
    & " From    MotivoNANCObj M JOIN OBJETO O ON O.COBJETOCOD = M.COBJETOCOD " _
    & " Where   nMotivocod =" & pnMotivoCod & ""

Set Rs = oCon.CargaRecordSet(Sql)
Set GetMotivosObjNivel = Rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetNroNotaCargoAbono(ByVal pnTpoDoc As TpoDoc) As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function

Sql = " Select Isnull(Max(cDocNro),0)  as MaxNota " _
    & " From NotaAbonoCargo " _
    & " Where  nDocTpo = '" & pnTpoDoc & "'"

Set Rs = oCon.CargaRecordSet(Sql)
If Not Rs.EOF And Not Rs.BOF Then
    GetNroNotaCargoAbono = Format(Rs!MaxNota + 1, "00000000")
Else
    GetNroNotaCargoAbono = Format(Rs!MaxNota + 1, "00000000")
End If
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetDetalleObjetos(ByVal pnObjetoCod As TpoObjetos, Optional ByVal psFiltro As String = "", Optional pnNivel As MuestraIF = MuestraCuentas) As ADODB.Recordset
Dim rs1 As New ADODB.Recordset
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
    Case ObjPersona
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function
Sql = " Select  M.cObjetoCod, M.cFiltro " _
    & " From    MotivoNANCObj M " _
    & " Where   M.nMotivocod =" & pnMotivoCod & " and cObjetoCod='" & psCodObjeto & "'"
GetFiltroMotivoObj = ""
Set Rs = oCon.CargaRecordSet(Sql)
If Not Rs.EOF And Not Rs.BOF Then
    GetFiltroMotivoObj = Rs!cFiltro
End If
Rs.Close
Set Rs = Nothing

oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetChequesNoDepositados(ByVal pnMoneda As Moneda, Optional ByVal psAreaCod As String = "", Optional psAgeCod As String = "") As ADODB.Recordset
Dim Sql As String
Dim Rs As New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltro As String


Set oCon = New COMConecta.DCOMConecta

If oCon.AbreConexion = False Then Exit Function

lsFiltro = ""
If psAreaCod <> "" Or psAgeCod <> "" Then
    lsFiltro = "  AND DR.CAREACOD ='" & psAreaCod & "' AND DR.CAGECOD='" & psAgeCod & "' "
End If

Sql = "Select   DISTINCT 0 as OK, P.cPersNombre AS Banco, DR.cNroDoc,DR.cIFCta,  " _
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

Set Rs = oCon.CargaRecordSet(Sql)
Set GetChequesNoDepositados = Rs
oCon.CierraConexion
Set oCon = Nothing
End Function





