VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NDocRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsServerCom As String
Dim vsServerPers As String
Public Function RegistroChequesNegocio(ByVal psMovNro As String, ByVal psOpeCod As String, _
                                ByVal psMovDesc As String, ByVal psTipoIF As String, _
                                ByVal psNumCheque As String, ByVal psPersCodIF As String, ByVal pnPlaza As ChequePlaza, _
                                ByVal psCtaCheque As String, ByVal pnImporte As Currency, ByVal pdFechaReg As Date, _
                                ByVal pdFechaVal As Date, ByVal psFormatoFecha As String, ByVal pnMonedaCheque As Moneda, _
                                Optional ByVal pnEstadoCheque As ChequeEstado = gChqEstEnValorizacion, _
                                Optional ByVal pnEstadoCG As CGEstadosChq = gCGEstadosChqRecibido, _
                                Optional ByVal pnConfirmaCaja As CGEstadoConfCheque = ChqCGSinConfirmacion, _
                                Optional psAreaCod As String, Optional psAgeCod As String) As Integer
Dim oMov As DMov
Dim lbTrans As Boolean
Dim lnMovNro As Long

On Error GoTo RegistroChequesNegocioErr
Set oMov = New DMov
oMov.BeginTrans
lbTrans = True
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)
If psNumCheque <> "" Then
    oMov.InsertaMovDoc lnMovNro, TpoDocCheque, psNumCheque, Format(pdFechaReg, gsFormatoFecha)
End If
'grabacion dentro de la tabla de cheques
oMov.InsertaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pnPlaza, psCtaCheque, pnImporte, pdFechaReg, _
                        pdFechaVal, pnEstadoCG, pnConfirmaCaja, pnMonedaCheque, , psAreaCod, psAgeCod
oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pdFechaReg, gChqEstEnValorizacion, psMovNro, psCtaCheque

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

Dim oMov As DMov
Dim lsSubCuenta As String
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New DMov

On Error GoTo RegistroChequesContabErr
lsSubCuenta = ""
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

Dim lsCtaCaracterOut As String
If oMov.CuentaEsPendiente(psCtaHaber + lsSubCta, lsCtaCaracterOut, "A") Then
    If lsCtaCaracterOut = "A" Then
        oMov.InsertaMovPendientesRend lnMovNro, psCtaHaber + lsSubCta, pnImporte * -1
    End If
End If
If Not rsObjHaber Is Nothing Then
    If rsObjHaber.State = adStateOpen Then
        rsObjHaber.MoveFirst
        Do While Not rsObjHaber.EOF
            lnMovOrden = lnMovOrden + 1
            Select Case rsObjHaber!cObjetoCod
                Case ObjCMACAgencias
                    oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!cObjetoCod
                    oMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!objeto, ""
                Case ObjCMACArea
                    oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rsObjHaber!cObjetoCod
                    oMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, "", rsObjHaber!objeto
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
    oMov.InsertaMovDoc lnMovNro, TpoDocCheque, psNumCheque, Format(pdFechaReg, gsFormatoFecha)
End If
'grabacion dentro de la tabla de cheques
oMov.InsertaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pnPlaza, psCtaCheque, pnImporte, pdFechaReg, _
                        pdFechaVal, pnEstadoCG, pnConfirmaCaja, pnMonedaCheque, , psAreaCodChq, psAgeCodChq

oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pdFechaReg, gChqEstEnValorizacion, psMovNro, psCtaCheque
If pnMovRef <> 0 Then
    oMov.InsertaMovRef lnMovNro, pnMovRef
End If
If Mid(psOpeCod, 3, 1) = "2" Then
    oMov.GeneraMovME lnMovNro, psMovNro
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
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Set oConect = New DConecta
Set rs = New ADODB.Recordset

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
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oIni = New ClasIni

vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing

End Sub
Public Function GetChequesValidos(ByVal dFecha As Date) As Recordset
Dim sSql As String
Dim clsConecta As DConecta
sSql = "SELECT  D.cNroDoc, P.cPersNombre, P.cPersCod  " _
    & "FROM     DocRec D INNER JOIN DocRecEst E INNER JOIN " _
    & vsServerPers & "Persona P ON E.cPersCod = E.cPersCod ON D.nTpoDoc = E.nTpoDoc AND D.cNroDoc = " _
    & "         E.cNroDoc AND D.cPersCod = E.cPersCod AND D.cIFTpo = E.cIFTpo " _
    & " WHERE D.nEstado = " & gChqEstEnValorizacion & " " _
    & "         AND E.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' " _
    & "         AND E.cMovNro IN (Select MAX(cMovNro) FROM DocRecEst E1  " _
    & "                           WHERE E1.nTpoDoc = E.nTpoDoc AND E1.cNroDoc = E.cNroDoc AND E1.cPersCod = E.cPersCod AND " _
    & "                           E1.cIFTpo = E.cIFTpo)"
Set clsConecta = New DConecta
If clsConecta.AbreConexion = False Then Exit Function
Set GetChequesValidos = clsConecta.CargaRecordSet(sSql)
clsConecta.CierraConexion
Set clsConecta = Nothing
End Function
Public Function GetCheques(ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psAgeCod As String, Optional psNumCheque As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim lsFiltro As String


Set oCon = New DConecta
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
Public Function GetPersCuentaAho(ByVal psCodCta As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim lsFiltro As String


Set oCon = New DConecta
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


Public Function GetEstadosCheques(ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta

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
                                    ByVal psCodCtaAho As String, ByVal pdFecsis As Date, Optional psCtaCheque As String = "") As Integer
Dim oMov As DMov
Dim lbTrans As Boolean
Dim lnMovNro As Long

On Error GoTo GrabaCambioEstadoChequeErr
Set oMov = New DMov
GrabaCambioEstadoCheque = 1
oMov.BeginTrans
lbTrans = True
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovOpeVarias lnMovNro, psNumCheque, Left(psOpeDesc, 12), pnImporte
Select Case psOpeCod
    Case gOpeChequesAnulacion, gOpeChequesRechazo
        oMov.ActualizaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, , , , psCtaCheque, , , , , pnEstadoCheque
        oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pdFecsis, pnEstadoCheque, psMovNro, psCtaCheque
    Case gOpeChequesValorización
        oMov.ActualizaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, , , , psCtaCheque, , , pnConfirmaCaja, , pnEstadoCheque
        oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pdFecsis, pnEstadoCheque, psMovNro, psCtaCheque
        'falta actualizar el saldo de la cuenta de ahorros que afecto el cheque
        If psCodCtaAho <> "" Then
        End If
    Case gOpeChequesModFecVal
        oMov.ActualizaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, , pdFechaVal, , psCtaCheque, , , , , pnEstadoCheque
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
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim lsFiltro As String


Set oCon = New DConecta
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
Public Function GrabaExtornoCheque(ByVal pnMovNroAnt As Long, _
                                   ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psOpeDesc As String, _
                                    ByVal psMovDesc As String, _
                                    ByVal psNumCheque As String, ByVal psPersCodIF As String, _
                                    ByVal psTipoIF As String, _
                                    ByVal pnImporte As Currency, Optional pnEstadoCheque As ChequeEstado, _
                                    Optional psCodCtaAho As String, Optional pdFecsis As Date, Optional pbEliminaMov As Boolean = True, Optional psCtaIFChq As String = "") As Integer
Dim oMov     As DMov
Dim lbTrans  As Boolean
Dim lnMovNro As Long
Dim lsMsgErr As String

On Error GoTo GrabaExtornoChequeErr
Set oMov = New DMov
GrabaExtornoCheque = 1
oMov.BeginTrans
lbTrans = True
oMov.ExtornaMovimiento psMovNro, pnMovNroAnt, psOpeCod, psMovDesc, pbEliminaMov, psMovNro
lnMovNro = oMov.GetnMovNro(psMovNro)
Select Case psOpeCod
    Case gOpeChequesExtAnulación, gOpeChequesValorización
        oMov.InsertaMovOpeVarias lnMovNro, psNumCheque, Left(psOpeDesc, 12), pnImporte
        oMov.ActualizaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, , , , psCtaIFChq, , , , , pnEstadoCheque
        oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pdFecsis, pnEstadoCheque, psMovNro, psCtaIFChq
    Case gOpeChequesExtValorización
        oMov.InsertaMovOpeVarias lnMovNro, psNumCheque, Left(psOpeDesc, 12), pnImporte
        oMov.ActualizaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, , , , psCtaIFChq, , , , , pnEstadoCheque
        oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pdFecsis, pnEstadoCheque, psMovNro, psCtaIFChq
        'Falta actualizar el saldo de la cuenta de ahorros que afecto el cheque
        If psCodCtaAho <> "" Then
        End If
        
    'Ope de Caja General
    Case gOpeCGExtBcoDepCheques, gOpeCGExtBcoDepChequesME
        oMov.ActualizaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, , , , psCtaIFChq, , gCGEstadosChqRecibido, ChqCGConfirmado
    Case gOpeCGExtBcoRegCheques, gOpeCGExtBcoRegChequesME, _
         gOpeCGExtBcoRecepChqRegAgencias, gOpeCGExtBcoRecepChqRegAgenciasME
        oMov.ActualizaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, , , , psCtaIFChq, , , , , gsChqEstExtornado
        oMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pdFecsis, gsChqEstExtornado, psMovNro, psCtaIFChq
End Select
GrabaExtornoCheque = 0
oMov.CommitTrans
lbTrans = False
Exit Function
GrabaExtornoChequeErr:
lsMsgErr = Err.Description
    If lbTrans Then
        lbTrans = False
        oMov.RollbackTrans
    End If
    Err.Raise Err.Number, Err.Source, lsMsgErr
End Function
Public Sub RegistroNotasAbonoCargo(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pnEstado As NotaCargoAbonoEstado, _
                                    ByVal pnMotivo As Integer, ByVal pnMonto As Currency, ByVal psMovNro As String, _
                                    Optional psObjetoCod As String = "", Optional psObjeto As String = "", Optional pnMoneda As Moneda = gMonedaNacional)
        
Dim oMov As DMov
Dim lbTrans As Boolean
On Error GoTo RegistroNotasAbonoCargoErr
Set oMov = New DMov
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
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsFiltro As String
Set rs = New ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
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
Public Function GetNotasCargoAbonoEst(ByVal pnTpoDoc As TpoDoc, Optional ByVal pnEstado As NotaCargoAbonoEstado = -1, Optional ByVal pnMoneda As Moneda = gMonedaNacional) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim oCon As DConecta
Dim lsFiltroEst As String
Set oCon = New DConecta
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
Public Function GetDatosNotaAC(ByVal pnTpoDoc As TpoDoc, ByVal psNroNANC As String, Optional ByVal pnEstado As NotaCargoAbonoEstado = -1) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim oCon As DConecta
Dim lsFiltroEst As String
Set oCon = New DConecta
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

Public Function GetMotivosObjNivel(ByVal pnMotivoCod As Long) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta

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
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta

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
Public Function GetDetalleObjetos(ByVal pnObjetoCod As TpoObjetos, Optional ByVal psFiltro As String = "", Optional pnNivel As MuestraIF = MuestraCuentas) As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim oRHAreas As DActualizaDatosArea
Dim oContFunct As NContFunciones
Dim oCtaIf As NCajaCtaIF
Dim oEfect As Defectivo
Dim lsFiltro As String

Set oCtaIf = New NCajaCtaIF
Set oContFunct = New NContFunciones
Set oRHAreas = New DActualizaDatosArea
Set oEfect = New Defectivo

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
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta

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
Public Function GetChequesNoDepositados(ByVal pnMoneda As Moneda, Optional ByVal psAreaCod As String = "", Optional psAgeCod As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim lsFiltro As String


Set oCon = New DConecta

If oCon.AbreConexion = False Then Exit Function

lsFiltro = ""
If psAreaCod <> "" Or psAgeCod <> "" Then
    lsFiltro = "  AND DR.CAREACOD ='" & psAreaCod & "' AND DR.CAGECOD='" & psAgeCod & "' "
End If

sql = "Select   DISTINCT 0 as OK, P.cPersNombre AS Banco, DR.cNroDoc,DR.cIFCta,  " _
    & "         CONVERT(VARCHAR(10),DR.dValorizacion,103) AS FECHA, DR.nMonto, " _
    & "         ISNULL(C.cConsDescripcion,'') AS EstActual, ISNULL(DRC.cCtaCod,'') AS CuentaAhorros, " _
    & "         ISNULL(AG.cAgeDescripcion,A.cAreaDescripcion) AS AreaAgencia, DR.cAreaCod,  DR.cAgeCod, " _
    & "         CASE WHEN DRC.cCtaCod IS NOT NULL THEN O.COBJETOCOD ELSE ISNULL(OBJ.COBJETOCOD,'6002' + convert(char(3),DR.nProducto)) END AS OBJETO, M.NMOVNRO, DR.cPersCod, DR.cIFTpo " _
    & " From    DOCREC DR " _
    & "         JOIN MOVDOC MD ON  MD.nDocTpo= DR.nTpoDoc  AND DR.cNroDoc=MD.cDocNro "
    
sql = sql & "   JOIN MOV    M  ON M.NMOVNRO = MD.NMOVNRO " _
    & "         JOIN DocRecEst dre ON dre.nTpoDoc = dr.nTpoDoc and dre.cNroDoc = dr.cNroDoc and dre.cPersCod = dr.cPersCod and dre.cIFCta = dr.cIFCta and dre.cMovNro = m.cMovNro and dre.nEstado in (0,1) " _
    & "         LEFT JOIN ( SELECT  MO.nMovNro, MO.COBJETOCOD " _
    & "                     FROM    MOVOBJ MO " _
    & "                     WHERE   MO.COBJETOCOD  LIKE '60%') AS OBJ " _
    & "         ON OBJ.NMOVNRO =M.NMOVNRO " _
    & "         JOIN PERSONA P          ON P.cPersCod =DR.cPersCod " _
    & "         LEFT JOIN CONSTANTE C       ON C.nConsValor= DR.nEstado AND C.nConsCod =1007 " _
    & "         LEFT JOIN (SELECT nTpoDoc, cNroDoc, cPersCod, cIFTpo, Max(cCtaCod) cCtaCod FROM DocRecCapta GROUP BY nTpoDoc, cNroDoc, cPersCod, cIFTpo ) DRC   ON DRC.nTpoDoc=DR.nTpoDoc AND DRC.cNrodoc=DR.cNrodoc " _
    & "                                         AND DRC.cPersCod=DR.cPersCod AND DRC.cIFTpo=DR.cIFTpo " _
    & "         JOIN AREAS A ON A.CAREACOD =DR.CAREACOD " _
    & "         LEFT JOIN AGENCIAS AG ON AG.CAGECOD =DR.CAGECOD " _
    & "         LEFT JOIN OBJETO O ON SUBSTRING(O.COBJETOCOD,5,3)=SUBSTRING(DRC.CCTACOD,6,3) " _
    & " WHERE   DR.cDepIF =" & gCGEstadosChqRecibido & " AND DR.nConfCaja=" & ChqCGConfirmado & " AND " _
    & "         DR.nEstado NOT IN (" & gChqEstAnulado & "," & gChqEstRechazado & "," & gsChqEstExtornado & ") " _
    & "         AND m.cOpecod = '900031' AND DR.nMoneda=" & pnMoneda & " AND M.NMOVFLAG=" & gMovFlagVigente _
    & "   and NOT EXISTS (SELECT mr.nMovNro FROM MovRef mr Join mov m1 on m1.nmovnro = mr.nmovnro " _
    & "                   WHERE RTRIM(ISNULL(mr.cAgeCodRef,'')) = '' and m1.nMovEstado = " & gMovEstContabMovContable & " and m1.nMovFlag = " & gMovFlagVigente & " and mr.nMovNroRef = m.nMovNro ) " _
    & lsFiltro

Set rs = oCon.CargaRecordSet(sql)
Set GetChequesNoDepositados = rs
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function CargaPendientesNegocioBancos(ByVal pnMoneda As Moneda, ByVal pbDepRet As Boolean, Optional psAgeCod As String = "", Optional bIncluirMonto As Boolean = False, Optional pnMonto As Double = 0) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim lsFiltro As String
Dim lsFiltroOpe As String
Dim lsDebeHaber As String
Set oCon = New DConecta
Dim lcOpeCod1, lcOpeCod2 As String '***Agregado por ELRO el 20120929, según OYP-RFC106-2012
Dim lbTodo, lbDeposito As Boolean '***Agregado por ELRO el 20120929, según OYP-RFC106-2012
If oCon.AbreConexion = False Then Exit Function

lsFiltro = ""
lsFiltroOpe = ""
lbTodo = True '***Agregado por ELRO el 20120929, según OYP-RFC106-2012

If psAgeCod <> "" Then
    lsFiltro = "  and  substring(m.cMovNro,18,2) ='" & psAgeCod & "' "
    lbTodo = False '***Agregado por ELRO el 20120929, según OYP-RFC106-2012
End If

If pbDepRet = True Then
   lsFiltroOpe = " and m.cOpeCod Between '300593' and '300599' " 'Operaciones Negocio Retiro
   '***Modificado por ELRO el 20120929, según OYP-RFC106-2012
   'lsDebeHaber = "'D'"
   lsDebeHaber = "D"
   '***Fin Modificado por ELRO el 20120929*******************
   lcOpeCod1 = "300593" '***Agregado por ELRO el 20120929, según OYP-RFC106-2012
   lcOpeCod2 = "300599" '***Agregado por ELRO el 20120929, según OYP-RFC106-2012
   lbDeposito = True '***Agregado por ELRO el 20120929, según OYP-RFC106-2012
Else
   lsFiltroOpe = " and m.cOpeCod Between '300493' and '300498' " 'Operaciones Negocio Deposito
   '***Modificado por ELRO el 20120929, según OYP-RFC106-2012
   'lsDebeHaber = "'H'"
   lsDebeHaber = "H"
   '***Fin Modificado por ELRO el 20120929*******************
   lcOpeCod1 = "300493" '***Agregado por ELRO el 20120929, según OYP-RFC106-2012
   lcOpeCod2 = "300498" '***Agregado por ELRO el 20120929, según OYP-RFC106-2012
End If
'***Modificado por ELRO el 20120929, según OYP-RFC106-2012
'sql = " Select Distinct m.nmovnro,substring(m.cMovNro,1,8),m.cOpeCod,o.cOpeDesc,mo.nMovImporte,mo.cNroDoc,mo.cReferencia,P.cPersNombre ,md.nDocTpo,d.cDocDesc,  " _
'      & " md.cDocNro , md.dDocFecha ,ag.cAgeDescripcion,m.cMovDesc,op.cCtaContcod ,ag.cAgecod   " _
'      & "  from  Mov m  " _
'      & " inner join MovOpeVarias mo ON mo.nmovnro =m.nmovnro  " _
'      & " inner join MovDoc md ON md.nmovnro =m.nmovnro  " _
'      & " inner join Documento d ON d.nDocTpo=md.nDoctpo  " _
'      & " inner join Persona p ON p.cPerscod =substring(mo.cReferencia,4,13) " _
'      & " inner join OpeTpo o ON o.cOpecod =m.cOpeCod  " _
'      & " inner join OpeCtaNeg op ON op.cOpeCod =o.cOpeCod  " _
'      & " inner join Agencias ag ON ag.cAgeCod = substring(m.cMovNro,18,2) " _
'      & " where m.nmovflag not in (1,2,3) and op.cOpeCtaDH = " & lsDebeHaber _
'      & " and NOT EXISTS (SELECT mr.nMovNro FROM MovRef mr Join mov m1 on m1.nmovnro = mr.nmovnro WHERE RTRIM(ISNULL(mr.cAgeCodRef,'')) = '' and (m1.nMovEstado = 10 or m1.nMovEstado = 13) and m1.nMovFlag = 0   and mr.nMovNroRef = m.nMovNro) " _
'      & " and  mo.nMoneda= " & pnMoneda & lsFiltroOpe _
'      & lsFiltro
'
'    If bIncluirMonto = True Then
'        sql = sql & "and mo.nMovImporte = " & pnMonto & " "
'    End If
'    sql = sql & "order by substring(m.cMovNro,1,8)"
sql = "exec stp_sel_devolverPendientesNegocioBancos " & IIf(lbTodo, 1, 0) & ", '" & psAgeCod & "', '" & lsDebeHaber & "', '" & lcOpeCod1 & "', '" & lcOpeCod2 & "', " & IIf(lbDeposito, 1, 0) & ", " & pnMoneda & " "
'***Fin Modificado por ELRO el 20120929*******************

Set rs = oCon.CargaRecordSet(sql)
Set CargaPendientesNegocioBancos = rs
oCon.CierraConexion
Set oCon = Nothing
End Function



Public Function GetImporteChequesCartera(ByVal pnMoneda As Moneda, pdFechaAl As Date, Optional ByRef lnCantChq As Long = 0) As Currency
Dim lnImporte As Currency
Dim rs As ADODB.Recordset
Set rs = GetResumenCheques(pnMoneda, gCGEstadosChqRecibido, pdFechaAl)
lnCantChq = 0
lnImporte = 0
Do While Not rs.EOF
   lnImporte = lnImporte + rs!nMonto
   lnCantChq = lnCantChq + rs!nCantidad
   rs.MoveNext
Loop
GetImporteChequesCartera = lnImporte
RSClose rs
End Function

'Public Function GetResumenCheques(ByVal pnMoneda As Moneda, pnEstadoChq As CGEstadosChq, pdFechaAl As Date) As ADODB.Recordset
'Dim sql As String
'Dim oCon As DConecta
'
'Set oCon = New DConecta
'If oCon.AbreConexion = False Then Exit Function
'
''sql = "SELECT   DR.cAreaCod,  DR.cAgeCod, ISNULL(AG.cAgeDescripcion,A.cAreaDescripcion) AS AreaAgencia, SUM(DR.nMonto) nMonto, COUNT(*) nCantidad " _
''    & " FROM    DOCREC DR JOIN MOVDOC MD ON  MD.nDocTpo= DR.nTpoDoc  AND DR.cNroDoc=MD.cDocNro " _
''    & "         JOIN MOV    M  ON M.NMOVNRO = MD.NMOVNRO " _
''    & "         JOIN AREAS A ON A.CAREACOD =DR.CAREACOD " _
''    & "         LEFT JOIN AGENCIAS AG ON AG.CAGECOD =DR.CAGECOD " _
''    & " WHERE   DR.cDepIF =" & gCGEstadosChqRecibido & " AND DR.nConfCaja=" & ChqCGConfirmado & " AND " _
''    & "         DR.nEstado NOT IN (" & gChqEstAnulado & "," & gChqEstRechazado & "," & gsChqEstExtornado & ") " _
''    & "         AND DR.nMoneda = " & pnMoneda & " AND M.NMOVFLAG=" & gMovFlagVigente & " " _
''    & "GROUP BY DR.cAreaCod, DR.cAgeCod, ISNULL(AG.cAgeDescripcion,A.cAreaDescripcion)"
'
'sql = " SELECT ISNULL(SUM(a.nMonto),0) nMonto, ISNULL(COUNT(*),0) nCantidad " _
'     & " FROM DocRec a" _
'     & " JOIN MOVDoc MD ON MD.nDocTpo = A.nTpoDoc And MD.cDocNro = A.cNroDoc" _
'     & " JOIN MOV M ON M.nMovNro = MD.nMovNro" _
'     & " WHERE  a.nMoneda = " & pnMoneda & " and M.cOpeCod IN ('" & gOpeCGOpeBancosRecibeChqAgMn & "','" & gOpeCGOpeBancosRecibeChqAgMe & "','" & gOpeCGOpeBancosDepChequesMN & "','" & gOpeCGOpeBancosDepChequesME & "') and M.nMovFlag not in (1) and datediff(dd, MD.dDocFecha, '" & Format(pdFechaAl, gsFormatoFecha) & "') = 0" _
'     & " and (a.cDepIF = '0'" _
'     & " or not exists" _
'     & "(select cmovnro from movref mr where mr.nMovNroRef = M.nMovNro" _
'     & "        and substring(M.cMovNro,1,8) < '" & Format(pdFechaAl, gsFormatoMovFecha) & "') )"
'
'Set GetResumenCheques = oCon.CargaRecordSet(sql)
'oCon.CierraConexion
'Set oCon = Nothing
'End Function

Public Function GetResumenCheques(ByVal pnMoneda As Moneda, pnEstadoChq As CGEstadosChq, pdFechaAl As Date) As ADODB.Recordset
Dim sql As String
Dim oCon As DConecta

Set oCon = New DConecta
If oCon.AbreConexion = False Then Exit Function

'sql = "SELECT   DR.cAreaCod,  DR.cAgeCod, ISNULL(AG.cAgeDescripcion,A.cAreaDescripcion) AS AreaAgencia, SUM(DR.nMonto) nMonto, COUNT(*) nCantidad " _
'    & " FROM    DOCREC DR JOIN MOVDOC MD ON  MD.nDocTpo= DR.nTpoDoc  AND DR.cNroDoc=MD.cDocNro " _
'    & "         JOIN MOV    M  ON M.NMOVNRO = MD.NMOVNRO " _
'    & "         JOIN AREAS A ON A.CAREACOD =DR.CAREACOD " _
'    & "         LEFT JOIN AGENCIAS AG ON AG.CAGECOD =DR.CAGECOD " _
'    & " WHERE   DR.cDepIF =" & gCGEstadosChqRecibido & " AND DR.nConfCaja=" & ChqCGConfirmado & " AND " _
'    & "         DR.nEstado NOT IN (" & gChqEstAnulado & "," & gChqEstRechazado & "," & gsChqEstExtornado & ") " _
'    & "         AND DR.nMoneda = " & pnMoneda & " AND M.NMOVFLAG=" & gMovFlagVigente & " " _
'    & "GROUP BY DR.cAreaCod, DR.cAgeCod, ISNULL(AG.cAgeDescripcion,A.cAreaDescripcion)"

sql = " SELECT ISNULL(SUM(a.nMonto),0) nMonto, ISNULL(COUNT(*),0) nCantidad " _
     & " FROM DocRec a" _
     & " JOIN MOVDoc MD ON MD.nDocTpo = A.nTpoDoc And MD.cDocNro = A.cNroDoc" _
     & " JOIN MOV M ON M.nMovNro = MD.nMovNro" _
     & " WHERE  a.nMoneda = " & pnMoneda & " and M.cOpeCod IN ('" & gOpeCGOpeBancosRecibeChqAgMn & "','" & gOpeCGOpeBancosRecibeChqAgMe & "','" & gOpeCGOpeBancosDepChequesMN & "','" & gOpeCGOpeBancosDepChequesME & "') and M.nMovFlag not in (1) and datediff(dd, MD.dDocFecha, '" & Format(pdFechaAl, gsFormatoFecha) & "') = 0" _
     & " and (a.cDepIF = '0'" _
     & " or not exists" _
     & "(select cmovnro from movref mr where mr.nMovNroRef = M.nMovNro" _
     & "        and substring(M.cMovNro,1,8) < '" & Format(pdFechaAl, gsFormatoMovFecha) & "') )"

Set GetResumenCheques = oCon.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing
End Function

'EJVG20121126 ***
Public Function grabarTalonarioCheque(ByVal psIFTpo As String, ByVal psPersCod As String, ByVal psCtaIFCod As String, ByVal pnNroChequeIni As Long, ByVal pnNroChequeFin As Long, ByVal psUser As String, ByVal pdFecha As Date) As String
    Dim oDocRec As New DDocRec
    Dim lbTrans As Boolean
    Dim i As Long
    Dim lnTalonarioId As Long
    
    On Error GoTo ErrgrabarTalonarioCheque
    grabarTalonarioCheque = ""
    oDocRec.dBeginTrans
    lbTrans = True
    
    lnTalonarioId = oDocRec.InsertaChequeTalonario(psIFTpo, psPersCod, psCtaIFCod, psUser, pdFecha)
    If lnTalonarioId > 0 Then
        For i = pnNroChequeIni To pnNroChequeFin
            Call oDocRec.InsertaChequeTalonarioDet(lnTalonarioId, i)
        Next
    Else
        grabarTalonarioCheque = "No se pudo registrar el talonario de cheques"
        oDocRec.dRollbackTrans
        Set oDocRec = Nothing
        Exit Function
    End If
    
    oDocRec.dCommitTrans
    lbTrans = False
    
    Set oDocRec = Nothing
    Exit Function
ErrgrabarTalonarioCheque:
    grabarTalonarioCheque = Err.Description
    If lbTrans Then
        oDocRec.dRollbackTrans
        Set oDocRec = Nothing
    End If
End Function
Public Function RecuperaChequesTalonario(ByVal psIFTpo As String, ByVal psPersCod As String, ByVal psCtaIFCod As String, ByVal pnNroChequeIni As Long, ByVal pnNroChequeFin As Long) As ADODB.Recordset
On Error GoTo ErrRecuperaChequesTalonario
    Dim oDocRec As New DDocRec
    Set RecuperaChequesTalonario = oDocRec.RecuperaChequesTalonario(psIFTpo, psPersCod, psCtaIFCod, pnNroChequeIni, pnNroChequeFin)
    Set oDocRec = Nothing
    Exit Function
ErrRecuperaChequesTalonario:
    Err.Raise Err.Number, "RecuperaChequesTalonario", Err.Description
End Function
Public Function RecuperaChequeExistenteEnMovDoc(ByVal psIFTpo As String, ByVal psPersCod As String, ByVal psCtaIFCod As String, ByVal psNroCheque As String) As ADODB.Recordset
On Error GoTo ErrRecuperaChequeExistenteEnMovDoc
    Dim oDocRec As New DDocRec
    Set RecuperaChequeExistenteEnMovDoc = oDocRec.RecuperaChequeMovDoc(psIFTpo, psPersCod, psCtaIFCod, psNroCheque)
    Set oDocRec = Nothing
    Exit Function
ErrRecuperaChequeExistenteEnMovDoc:
    Err.Raise Err.Number, "RecuperaChequeExistenteEnMovDoc", Err.Description
End Function
'END EJVG *******

'RIRO20131212 ERS137 *****
Public Function CargaPendientesTransf(ByVal psAgeCod As String, ByVal pnMoneda As Moneda, Optional ByVal pnMovNro As Long = 0) As ADODB.Recordset

    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lbTodo, lbDeposito As Boolean
    
    On Error GoTo ERROR
    
    If oCon.AbreConexion = False Then Exit Function
    sql = "exec stp_sel_devolverPendientesTransferencia '" & psAgeCod & "', 'H', " & pnMoneda & ", " & pnMovNro
    Set rs = oCon.CargaRecordSet(sql)
    Set CargaPendientesTransf = rs
    oCon.CierraConexion
    Set oCon = Nothing
    
    Exit Function
ERROR:
    Err.Raise Err.Number, "RecuperaChequeExistenteEnMovDoc", Err.Description
End Function
'FIN RIRO ******
'EJVG20140127 ***
Public Function ValorizaChequeNegocio(ByRef pMatDatos() As Variant, ByVal psOpeCod As String, ByVal pdFecha As Date, ByVal psAgeCod As String, ByVal psUserCod As String) As Boolean
On Error GoTo ErrValorizaChequeNegocio
    Dim oMov As New DMov
    Dim lsMovNro As String
    Dim lnMovNro As Long, lnID As Long
    Dim i As Long
    Dim bTrans As Boolean
    Dim lsIDs As String
    
    oMov.BeginTrans
    bTrans = True
    lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, psUserCod)
    For i = 1 To UBound(pMatDatos, 2)
        lnID = pMatDatos(1, i)
        oMov.ValorizaChequeNeg lnID, CDate(pdFecha & " " & Format(Time, "hh:mm:ss")), lsMovNro 'lsMovNro agregado PASI20140615
        lsIDs = lsIDs & CStr(lnID) & ","
    Next
    If Len(lsIDs) > 0 Then
        lsIDs = Left(lsIDs, Len(lsIDs) - 1)
    End If
    
    oMov.InsertaMov lsMovNro, psOpeCod, Left("Valorización de Cheques N° " & lsIDs, 300), gMovEstContabMovContable, gMovFlagVigente
    oMov.CommitTrans
    bTrans = False
    ValorizaChequeNegocio = True
    Set oMov = Nothing
    Exit Function
ErrValorizaChequeNegocio:
    ValorizaChequeNegocio = False
    If bTrans Then
        oMov.RollbackTrans
        Set oMov = Nothing
    End If
    Err.Raise Err.Number, "ValorizaChequeNegocio", Err.Description
End Function
Public Function RechazaChequeNegocio(ByRef pMatDatos() As Variant, ByVal psOpeCod As String, ByVal pdFecha As Date, ByVal psAgeCod As String, ByVal psUserCod As String, ByVal psGlosa As String, ByRef psMovNroImpre As String) As Boolean
On Error GoTo ErrRechazaChequeNegocio
    Dim oMov As New DMov
    Dim oRSDet As New ADODB.Recordset
    Dim lsMovNro As String
    Dim lnMovNro As Long, lnID As Long
    Dim i As Long
    Dim ldFechaRechaza As Date
    Dim bTrans As Boolean
    Dim lsCtaContD As String, lsCtaContH As String, lsMovNroAsiento As String
    Dim lsIFTpo As String, lsPersCod As String, lsCtaIFCod As String
    Dim lnMovNroRef As Long
    
    oMov.BeginTrans
    bTrans = True
    For i = 1 To UBound(pMatDatos, 2)
        lnID = pMatDatos(1, i)
        lsCtaContD = pMatDatos(2, i)
        lsCtaContH = pMatDatos(3, i)
        lsIFTpo = pMatDatos(4, i)
        lsPersCod = pMatDatos(5, i)
        lsCtaIFCod = pMatDatos(6, i)
        
        ldFechaRechaza = CDate(pdFecha & " " & Format(Time, "hh:mm:ss"))
        Set oRSDet = oMov.ChequeDetxID(lnID)
        oMov.RechazaChequeNeg lnID, ldFechaRechaza
        Do While Not oRSDet.EOF
            lnMovNroRef = oRSDet!nMovNro
            Sleep 1000
            lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, psUserCod)
            oMov.InsertaMov lsMovNro, psOpeCod, Left(psGlosa, 300), gMovEstContabMovContable, gMovFlagVigente
            lnMovNro = oMov.GetnMovNro(lsMovNro)
            oMov.InsertaChequeEstado oRSDet!nTpoDoc, oRSDet!cNroDoc, oRSDet!cPerscod, oRSDet!cIFTpo, ldFechaRechaza, gChqEstRechazado, lsMovNro, oRSDet!cIFCta
            oMov.InsertaMovCta lnMovNro, 1, lsCtaContD, Abs(oRSDet!nMonto)
            oMov.InsertaMovCta lnMovNro, 2, lsCtaContH, Abs(oRSDet!nMonto) * -1
            oMov.InsertaMovObj lnMovNro, 2, 1, ObjEntidadesFinancieras
            oMov.InsertaMovObjIF lnMovNro, 2, 1, lsPersCod, lsIFTpo, lsCtaIFCod
            oMov.InsertaMovRef lnMovNro, lnMovNroRef
            oMov.InsertaMovDoc lnMovNro, oRSDet!nTpoDoc, oRSDet!cNroDoc, Format(ldFechaRechaza, "yyyymmdd hh:mm:ss")
            oMov.InsertaMovDocRec lnMovNro, oRSDet!nTpoDoc, oRSDet!cNroDoc, oRSDet!cPerscod, oRSDet!cIFTpo, oRSDet!cIFCta
            oMov.ActualizaNroMovRechazo lnMovNroRef, lnMovNro
            
            If Mid(psOpeCod, 3, 1) = "2" Then
                oMov.GeneraMovME lnMovNro, lsMovNro
            End If
            oMov.ActualizaSaldoMovimiento lsMovNro, "+"
            lsMovNroAsiento = lsMovNroAsiento & lsMovNro & ","
            oRSDet.MoveNext
        Loop
    Next
    If Len(lsMovNroAsiento) > 0 Then
        lsMovNroAsiento = Mid(lsMovNroAsiento, 1, Len(lsMovNroAsiento) - 1)
    End If
    oMov.CommitTrans
    bTrans = False
    RechazaChequeNegocio = True
    psMovNroImpre = lsMovNroAsiento
    
    Set oRSDet = Nothing
    Set oMov = Nothing
    Exit Function
ErrRechazaChequeNegocio:
    RechazaChequeNegocio = False
    If bTrans Then
        oMov.RollbackTrans
        Set oMov = Nothing
    End If
    Err.Raise Err.Number, "RechazaChequeNegocio", Err.Description
End Function
Public Function ExtornaValorizaChequeNegocio(ByVal pnId As Long, ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psGlosa As String) As Boolean
On Error GoTo ErrExtornaValorizaChequeNegocio
    Dim oMov As New DMov
    Dim lsMovNro As String
    Dim bTrans As Boolean
    
    oMov.BeginTrans
    bTrans = True
    oMov.ExtornaValorizaChequeNeg pnId
    oMov.InsertaMov psMovNro, psOpeCod, Left("Extorno Valorización de Cheque N° " & pnId & " : " & psGlosa, 300), gMovEstContabNoContable, gMovFlagVigente
    oMov.CommitTrans
    bTrans = False
    ExtornaValorizaChequeNegocio = True
    Set oMov = Nothing
    Exit Function
ErrExtornaValorizaChequeNegocio:
    ExtornaValorizaChequeNegocio = False
    If bTrans Then
        oMov.RollbackTrans
        Set oMov = Nothing
    End If
    Err.Raise Err.Number, "ExtornaValorizaChequeNegocio", Err.Description
End Function
Public Function ExtornarRechazoCheque(ByVal pnId As Long, ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psGlosa As String) As Boolean
    Dim oMov As New DMov
    Dim lbTrans As Boolean
    Dim lnMovNro As Long
    Dim rsDet As New ADODB.Recordset

    On Error GoTo ErrExtornarRechazoCheque
    oMov.BeginTrans
    lbTrans = True
    Set rsDet = oMov.ChequeDetxID(pnId)
    oMov.InsertaMov psMovNro, psOpeCod, Left(psGlosa, 300), gMovEstContabNoContable, gMovFlagDeExtorno
    lnMovNro = oMov.GetnMovNro(psMovNro)
    Do While Not rsDet.EOF
        If rsDet!nMovNroRechazo > 0 Then
            oMov.ActualizaMov rsDet!nMovNroRechazo, , , gMovFlagExtornado
            oMov.InsertaMovRef lnMovNro, rsDet!nMovNroRechazo
        End If
        rsDet.MoveNext
    Loop
    oMov.ExtornaRechazoChequeNeg pnId
    oMov.CommitTrans
    lbTrans = False
    ExtornarRechazoCheque = True
    Set oMov = Nothing
    Exit Function
ErrExtornarRechazoCheque:
    ExtornarRechazoCheque = False
    If lbTrans Then
        oMov.RollbackTrans
        Set oMov = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'END EJVG *******
