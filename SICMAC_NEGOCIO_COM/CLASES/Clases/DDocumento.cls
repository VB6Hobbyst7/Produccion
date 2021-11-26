VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A80C4CD03B9"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
 Option Base 0
Option Explicit
Dim dbConec As DConecta
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
   Set dbConec = New DConecta
   If dbConec.AbreConexion(sConexion) Then
      Set CargaDocumento = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaDocumentoErr:
   Call RaiseError(MyUnhandledError, "DDocumento:CargaDocumento Method")
End Function

Public Sub InsertaDocumento(pnDocTpo As Long, psDocDesc As String, psDocAbrev As String, psMovNro As String)
   On Error GoTo InsertaDocumentoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion(sConexion) Then
      psSql = "INSERT " & sCentralCom & "Documento (nDocTpo, cDocDesc, cDocAbrev, cUltimaActualizacion) " _
         & "VALUES (" & pnDocTpo & ", '" & psDocDesc & "','" & psDocAbrev & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaDocumentoErr:
   Call RaiseError(MyUnhandledError, "DDocumento:InsertaDocumento Method")
End Sub

Public Sub ActualizaDocumento(pnDocTpo As Long, psDocDesc As String, psDocAbrev As String, psMovNro As String)
   On Error GoTo ActualizaDocumentoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion(sConexion) Then
      psSql = "UPDATE Documento Set cDocDesc = '" & psDocDesc & "', cDocAbrev = '" & psDocAbrev & "', cUltimaActualizacion = '" & psMovNro & "'  WHERE nDocTpo = " & pnDocTpo & " "
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizaDocumentoErr:
   Call RaiseError(MyUnhandledError, "DDocumento:ActualizaDocumento Method")
End Sub

Public Sub AsignaDocImpuesto(pnDocTpo As Long, psCtaContCod As String, psDocImpDH As String, psDocImpOpc As String, psMovNro As String)
   On Error GoTo AsignaDocImpuestoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion(sConexion) Then
      psSql = "INSERT " & sCentralCom & "DocImpuesto (nDocTpo, cCtaContCod, cDocImpDH, cDocImpOpc, cUltimaActualizacion) " _
         & "VALUES (" & pnDocTpo & ", '" & psCtaContCod & "','" & psDocImpDH & "','" & psDocImpOpc & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
AsignaDocImpuestoErr:
   Call RaiseError(MyUnhandledError, "DDocumento:AsignaDocImpuesto Method")
End Sub

Public Sub DesasignaDocImpuesto(pnDocTpo As Long, psCtaContCod As String)
   On Error GoTo DesasignaDocImpuestoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion(sConexion) Then
      psSql = "DELETE " & sCentralCom & "DocImpuesto WHERE nDocTpo = " & pnDocTpo & " and cCtaContCod = '" & psCtaContCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
DesasignaDocImpuestoErr:
   Call RaiseError(MyUnhandledError, "DDocumento:DesasignaDocImpuesto Method")
End Sub

Private Sub Class_Initialize()
Dim oIni As New ClasIni
sConexion = oIni.CadenaConexion
sCentralCom = oIni.BaseComunes
Set oIni = Nothing
End Sub
Public Function GetValidaDocProv(ByVal psPersCod As String, ByVal pnDocTpo As Long, ByVal psDocNro As String) As Boolean
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Set rs = New ADODB.Recordset
Set oConect = New DConecta
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
Dim oConect As DConecta
Dim sql As String
Set rs = New ADODB.Recordset
Set oConect = New DConecta
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
Dim oConect As DConecta
Dim sql As String
Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
    
    sSql = " Select  cNroDoc " _
        & " From    Docrec " _
        & " Where   nTpodoc = " & pnTpoDoc & " and cNroDoc = '" & psNroDoc & "' " _
        & "         and cPersCod = '" & psPersCod & "' and cIFtpo = '" & Format(psIFTpo, "00") & "'"
        
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
   Set dbConec = New DConecta
   If dbConec.AbreConexion(sConexion) Then
      lTrans = True
      dbConec.BeginTrans
      psSql = "DELETE DocImpuesto WHERE nDocTpo = " & pnDocTpo & " "
      dbConec.Ejecutar psSql
      psSql = "DELETE Documento WHERE nDocTpo = " & pnDocTpo & " "
      dbConec.Ejecutar psSql
      dbConec.CommitTrans
      lTrans = False
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaDocumentoErr:
   lsError = Err.Description
   If lTrans Then
      dbConec.RollbackTrans
   End If
   Err.Raise Err.Number + MyUnhandledError, "DDocumento:EliminaDocumento", lsError
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
       
   Set dbConec = New DConecta
   If dbConec.AbreConexion(sConexion) Then
      Set CargaDocImpuesto = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaDocImpuestoErr:
   Call RaiseError(MyUnhandledError, "DDocumento:CargaDocImpuesto Method")
End Function





