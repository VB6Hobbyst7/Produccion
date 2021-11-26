VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Dim DBConec As COMConecta.DCOMConecta
Dim oError As New COMConecta.COMErrorHandling
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sConexion As String
Dim sCentralCom As String
Public Function CargaDocumento(Optional pnDocTpo As Long = -1, Optional psDocAbrev As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
Dim sCond As String
   On Error GoTo CargaDocumentoErr
   If pnDocTpo <> -1 Then
      sCond = " WHERE nDocTpo = " & pnDocTpo & " "
   End If
   If psDocAbrev <> "" Then
      sCond = sCond & IIf(sCond = "", "WHERE", "AND") & " cDocAbrev = '" & psDocAbrev & "'"
   End If
   psSql = "SELECT nDocTpo, cDocDesc, cDocAbrev FROM " & sCentralCom & "Documento " & sCond
   Set DBConec = New COMConecta.DCOMConecta
   If DBConec.AbreConexion(sConexion) Then
      Set CargaDocumento = DBConec.CargaRecordSet(psSql, LockType)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaDocumentoErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DDocumento:CargaDocumento Method")
End Function

Public Sub InsertaDocumento(pnDocTpo As Long, psDocDesc As String, psDocAbrev As String, psMovNro As String)
   On Error GoTo InsertaDocumentoErr
   Set DBConec = New COMConecta.DCOMConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "INSERT " & sCentralCom & "Documento (nDocTpo, cDocDesc, cDocAbrev, cUltimaActualizacion) " _
         & "VALUES (" & pnDocTpo & ", '" & psDocDesc & "','" & psDocAbrev & "','" & psMovNro & "')"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
InsertaDocumentoErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DDocumento:InsertaDocumento Method")
End Sub

Public Sub ActualizaDocumento(pnDocTpo As Long, psDocDesc As String, psDocAbrev As String, psMovNro As String)
   On Error GoTo ActualizaDocumentoErr
   Set DBConec = New COMConecta.DCOMConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "UPDATE Documento Set cDocDesc = '" & psDocDesc & "', cDocAbrev = '" & psDocAbrev & "', cUltimaActualizacion = '" & psMovNro & "'  WHERE nDocTpo = " & pnDocTpo & " "
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
ActualizaDocumentoErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DDocumento:ActualizaDocumento Method")
End Sub

Public Sub AsignaDocImpuesto(pnDocTpo As Long, psCtaContCod As String, psDocImpDH As String, psDocImpOpc As String, psMovNro As String)
   On Error GoTo AsignaDocImpuestoErr
   Set DBConec = New COMConecta.DCOMConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "INSERT " & sCentralCom & "DocImpuesto (nDocTpo, cCtaContCod, cDocImpDH, cDocImpOpc, cUltimaActualizacion) " _
         & "VALUES (" & pnDocTpo & ", '" & psCtaContCod & "','" & psDocImpDH & "','" & psDocImpOpc & "','" & psMovNro & "')"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
AsignaDocImpuestoErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DDocumento:AsignaDocImpuesto Method")
End Sub

Public Sub DesasignaDocImpuesto(pnDocTpo As Long, psCtaContCod As String)
   On Error GoTo DesasignaDocImpuestoErr
   Set DBConec = New COMConecta.DCOMConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "DELETE " & sCentralCom & "DocImpuesto WHERE nDocTpo = " & pnDocTpo & " and cCtaContCod = '" & psCtaContCod & "'"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
DesasignaDocImpuestoErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DDocumento:DesasignaDocImpuesto Method")
End Sub

Private Sub Class_Initialize()
Dim oIni As New COMConecta.DCOMClasIni
sConexion = oIni.CadenaConexion
sCentralCom = oIni.BaseComunes
Set oIni = Nothing
End Sub
Public Function GetValidaDocProv(ByVal psPersCod As String, ByVal pnDocTpo As Long, ByVal psDocNro As String) As Boolean
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

sql = " SELECT  a.cMovNro " _
     & " FROM   MovGasto a JOIN MovDoc b ON b.cMovNro = a.cMovNro JOIN Mov m ON m.cMovNro = a.cMovNro " _
     & " WHERE  M.nMovFlag not in ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "')  and a.cPersCod = '" & psPersCod & "' and b.nDocTpo = " & pnDocTpo & " and b.cDocNro = '" & psDocNro & "'"

GetValidaDocProv = False
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetValidaDocProv = True
End If
rs.Close
Set rs = Nothing
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function VerificaDoc(ByVal pnDocTpo As TpoDoc, ByVal psNroDoc As String, Optional ByVal psPersCod As String = "") As Boolean
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta
Dim sql As String
Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
VerificaDoc = False
If oConect.AbreConexion = False Then Exit Function
    If pnDocTpo = TpoDocOrdenPago Then
        sSql = "SELECT a.nMovNro FROM  movdoc a " _
                & "   WHERE a.nDocTpo = " & pnDocTpo & " and a.cDocNro = '" & psNroDoc & "' "
    Else
        sSql = "SELECT a.nMovNro FROM  movdoc a JOIN MOVOBJIF b on a.nMovNro = b.nMovNro  " _
              & "WHERE a.nDocTpo = " & pnDocTpo & " and a.cDocNro = '" & psNroDoc & "' " _
              & "  and b.cPersCod ='" & psPersCod & "' "
    End If
    Set rs = oConect.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
       VerificaDoc = True
    End If
    rs.Close: Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Public Function VerificaCheque(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF) As Boolean
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta
Dim sql As String
Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
    
    sSql = " Select  cNroDoc " _
        & " From    Docrec " _
        & " Where   nTpodoc = " & pnTpoDoc & " and cNroDoc = '" & psNroDoc & "' " _
        & "         and cPersCod = '" & psPersCod & "' and cIFtpo = '" & Format(psIFTpo, "00") & "'"
        '& " and nEstado= '" & gChqEstRegistrado & "'"
        
    VerificaCheque = False
    Set rs = oConect.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
       VerificaCheque = True
    End If
    rs.Close: Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
End Function

'************************
'** NUEVAS FUNCIONES
'************************
Public Sub EliminaDocumento(pnDocTpo As Long)
Dim lTrans As Boolean
Dim lsError As String
   On Error GoTo EliminaDocumentoErr
   Set DBConec = New COMConecta.DCOMConecta
   If DBConec.AbreConexion(sConexion) Then
      lTrans = True
      DBConec.BeginTrans
      psSql = "DELETE DocImpuesto WHERE nDocTpo = " & pnDocTpo & " "
      DBConec.Ejecutar psSql
      psSql = "DELETE Documento WHERE nDocTpo = " & pnDocTpo & " "
      DBConec.Ejecutar psSql
      DBConec.CommitTrans
      lTrans = False
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
EliminaDocumentoErr:
   lsError = Err.Description
   If lTrans Then
      DBConec.RollbackTrans
   End If
   
End Sub

Public Function CargaDocImpuesto(pnDocTpo As Long, Optional psCtaCod As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
   On Error GoTo CargaDocImpuestoErr
   psSql = "SELECT d.cCtaContCod, c.cCtaContDesc, " & _
       "       CASE WHEN cDocImpDH ='D' THEN 'DEBE' " & _
       "            WHEN cDocImpDH ='H' THEN 'HABER' " & _
       "       END as Clase , " & _
       "       CASE WHEN cDocImpOpc ='1' THEN 'Obligatorio' " & _
       "            WHEN cDocImpOpc ='2' THEN 'Opcional' " & _
       "       END as Caracter , " _
       & "       d.cDocImpDH, d.cDocImpOpc, I.cImpAbrev,  I.nImpTasa, I.cImpDestino " _
       & " FROM  " & sCentralCom & "CtaCont c " _
       & "      JOIN " & sCentralCom & "DocImpuesto d ON d.cCtaContCod = c.cCtaContCod " _
       & "      JOIN " & sCentralCom & "Impuesto I ON I.cCtaContCod = D.cCtaContCod " _
       & IIf(pnDocTpo <> -1, "WHERE d.nDocTpo= " & pnDocTpo, "") & IIf(psCtaCod <> "", IIf(pnDocTpo <> -1, " WHERE ", " AND ") & "c.cCtaContCod = '" & psCtaCod & "'", "")
       
   Set DBConec = New COMConecta.DCOMConecta
   If DBConec.AbreConexion(sConexion) Then
      Set CargaDocImpuesto = DBConec.CargaRecordSet(psSql, LockType)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaDocImpuestoErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DDocumento:CargaDocImpuesto Method")
End Function

Public Function CargaDocVaucherCtaPendiente(ByVal dFechaI As Date, ByVal dFechaF As Date) As Recordset
   On Error GoTo CargaDocVaucherError
   
   psSql = "SELECT md.cDocNro,p.cPersNombre,md.dDocFecha,mc.nMovMonto,ch.cAreaCod,ch.cAgeCod,p.cPersCod,cCtaContCod,m.nMovNro" & _
           " FROM Mov m " & _
           " JOIN MOVCTA mct ON mct.nMovNro=m.nMovNro" & _
           " JOIN MOVCONT mc ON mc.nMovNro=m.nMovNro" & _
           " JOIN MOVDOC md ON md.nMovNro=m.nMovNro" & _
           " JOIN MOVGASTO mg ON mg.nMovNro=m.nMovNro" & _
           " LEFT JOIN MOVREF mr ON mr.nMovNro = m.nMovNro " & _
           " JOIN MOVCAJACHICA MCH ON M.nMovNro = MCH.nMovNro" & _
           " JOIN CAJACHICA CH ON CH.cAreaCod= MCH.cAreaCod " & _
           " and CH.cAgeCod= MCH.cAgeCod " & _
           " AND CH.nProcNro = MCH.nProcNro" & _
           " JOIN PERSONA p ON p.cPersCod =mg.cPersCod" & _
           " WHERE m.cOpeCod ='401324' and dDocFecha >= '" & Format(dFechaI, "yyyymmdd") & "' and dDocFecha <= '" & Format(dFechaF, "yyyymmdd") & "' "
  psSql = psSql & " AND cCtaContCod like '19111909%'  And m.nMovFlag = 0 " & _
           " AND Not Exists (Select MR.nMovNroRef FROM MOVREF MR " & _
           " Inner Join Mov MMR On MMR.nMovNro = MR.nMovNro" & _
           " Where MMR.nMovFlag = 0 And MR.nMovNroRef = M.nMovNro ) " & _
           " AND  m.nMovNro Not In (Select nMovNroRef FROM MOVREF)"
   psSql = psSql & " UNION " & _
           " SELECT md.cDocNro,p.cPersNombre,md.dDocFecha,mc.nMovMonto,ch.cAreaCod, CH.cAgeCod , " & _
           " p.cPersCod , cCtaContCod, mr.nMovNroref" & _
           " FROM Mov m" & _
           " JOIN MOVCTA mct ON mct.nMovNro=m.nMovNro" & _
           " JOIN MOVCONT mc ON mc.nMovNro=m.nMovNro" & _
           " JOIN MOVDOC md ON md.nMovNro=m.nMovNro" & _
           " JOIN MOVGASTO mg ON mg.nMovNro=m.nMovNro" & _
           " JOIN MOVREF mr ON mr.nMovNro=m.nMovNro JOIN MOV ref on ref.nMovNro=mr.nMovNroRef" & _
           " JOIN MOVCAJACHICA MCH ON M.nMovNro=MCH.nMovNro" & _
           " JOIN CAJACHICA CH ON CH.cAreaCod= MCH.cAreaCod  and CH.cAgeCod= MCH.cAgeCod  AND CH.nProcNro = MCH.nProcNro" & _
           " JOIN PERSONA p ON p.cPersCod =mg.cPersCod" & _
           " WHERE m.cOpeCod ='300591'  and dDocFecha >= '" & Format(dFechaI, "yyyymmdd") & "' and dDocFecha <= '" & Format(dFechaF, "yyyymmdd") & "' AND cCtaContCod like '19111909%'" & _
           " and ref.nMovFlag = 0 and m.nMovFlag = 1" & _
           " and mr.nMovNro =(Select Max (nMovNro)FROM MovRef where nMovNroref =mr.nMovNroref)"
           
   Set DBConec = New COMConecta.DCOMConecta
   If DBConec.AbreConexion(sConexion) Then
      Set CargaDocVaucherCtaPendiente = DBConec.CargaRecordSet(psSql)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaDocVaucherError:
   Call oError.RaiseError(oError.MyUnhandledError, "DDocumento:CargaDocImpuesto Method")
End Function


Public Function GetDatosCajaChica(ByVal psAreaCh As String, ByVal psAgeCh As String, Optional ByVal DatosCajaChica As gDatosCajaChica) As Currency
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function

sql = "SELECT  cAreaCod, cAgeCod, nProcNro, nMontoAsig, nTopeEgresos,nMontoDesem,nSaldoAnt,nSaldo " _
    & " FROM   CAJACHICA  C " _
    & " WHERE  C.nProcNro = (SELECT MAX(nProcNro) " _
    & "                      FROM CAJACHICA C1 " _
    & "                      WHERE C1.cAreaCod = C.cAreaCod and C1.cAgeCod = C.cAgeCod ) " _
    & "        And cAreaCod='" & psAreaCh & "' and cAgeCod= '" & psAgeCh & "'"
GetDatosCajaChica = 0

Set rs = oCon.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    Select Case DatosCajaChica
        Case gMontoAsig
            GetDatosCajaChica = rs!nMontoAsig
        Case gMontoDesemb
            GetDatosCajaChica = rs!nMontoDesem
        Case gMontoTope
            GetDatosCajaChica = rs!nTopeEgresos
        Case gSaldoActual
            GetDatosCajaChica = rs!nSaldo
        Case gSaldoAnterior
            GetDatosCajaChica = rs!nSaldoAnt
        Case gNroCajaChica
            GetDatosCajaChica = rs!nProcNro

    End Select

End If

rs.Close
Set rs = Nothing
    oCon.CierraConexion
Set oCon = Nothing

End Function

Public Function CargaDocVaucherCtaPendienteExtorno(ByVal pdFecha As Date) As Recordset
   On Error GoTo CargaDocVaucherError
   
   psSql = " SELECT md.cDocNro,p.cPersNombre,md.dDocFecha,mc.nMovMonto,ch.cAreaCod," & _
           " CH.cAgeCod , p.cPersCod, cCtaContCod, m.nMovNro " & _
           " FROM Mov m  JOIN MOVCTA mct ON mct.nMovNro=m.nMovNro" & _
           " JOIN MOVCONT mc ON mc.nMovNro=m.nMovNro" & _
           " JOIN MOVDOC md ON md.nMovNro=m.nMovNro" & _
           " JOIN MOVGASTO mg ON mg.nMovNro=m.nMovNro" & _
           " LEFT JOIN MOVREF mr   ON mr.nMovNro=m.nMovNro" & _
           " JOIN MOVCAJACHICA MCH ON M.nMovNro=MCH.nMovNro" & _
           " JOIN CAJACHICA CH ON CH.cAreaCod= MCH.cAreaCod  and CH.cAgeCod= MCH.cAgeCod  AND CH.nProcNro = MCH.nProcNro" & _
           "  JOIN PERSONA p ON p.cPersCod =mg.cPersCod" & _
           " WHERE m.cOpeCod ='300591'  And m.nMovFlag = 0 AND cCtaContCod like '19111909%'" & _
           " and md.dDocFecha ='" & Format(pdFecha, "yyyymmdd") & "'" & _
           " AND Not Exists (Select MR.nMovNroRef" & _
           " FROM MOVREF MR  Inner Join Mov MMR On MMR.nMovNro = MR.nMovNro" & _
           " Where MMR.nMovFlag = 0 And MR.nMovNroRef = M.nMovNro )" & _
           " AND  m.nMovNro Not In (Select nMovNroRef FROM MOVREF) "

   Set DBConec = New COMConecta.DCOMConecta
   If DBConec.AbreConexion(sConexion) Then
      Set CargaDocVaucherCtaPendienteExtorno = DBConec.CargaRecordSet(psSql)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaDocVaucherError:
   Call oError.RaiseError(oError.MyUnhandledError, "DDocumento:CargaDocImpuesto Method")
End Function

