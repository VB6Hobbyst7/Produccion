VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim dbConec As COMConecta.DCOMConecta
Dim psSql As String
Dim prs As ADODB.Recordset

Dim sConexion As String
Dim sCentralCom As String

Public Function CargaOpeCta(psOpeCod As String, Optional ByVal psDH As String = "", Optional ByVal psOpeCtaOrden As String = "", Optional LockType As LockTypeEnum = adLockReadOnly, Optional pbGetClase As Boolean = False) As Recordset
   On Error GoTo CargaOpeCtaErr
   Set dbConec = New COMConecta.DCOMConecta
   Dim lsFiltroDH As String
   Dim lsFiltroOrden As String
    If psDH <> "" Then
        lsFiltroDH = " AND a.cOpeCtaDH ='" & psDH & "' "
    End If
    If psOpeCtaOrden <> "" Then
        lsFiltroOrden = " And a.cOpeCtaOrden='" & psOpeCtaOrden & "' "
    End If
   If dbConec.AbreConexion() Then
      psSql = "select " _
         & "  a.cCtaContCod, b.cCtaContDesc,   1 as Nivel ,  " _
         & "  CASE WHEN a.cOpeCtaDH = '" & gsContHaber & "' THEN '" & gsContHaberDesc & "' " _
         & "       WHEN a.cOpeCtaDH = '" & gsContDebe & "' THEN '" & gsContDebeDesc & "' END  as Clase, " _
         & "  CASE WHEN a.cOpeCtaOpc = '" & gOpeCtaCaracObligatorio & "' then '" & gsOpeCtaCaracterObligaDesc & "' WHEN a.cOpeCtaOpc='" & gOpeCtaCaracObligatorio & "' then '" & gsOpeCtaCaracterOpcionDesc & "' END as Caracter, " _
         & "  a.cOpeCod, a.cOpeCtaDH, a.cOpeCtaOpc, a.cCtaContN, a.cOpeCtaOrden " & IIf(pbGetClase, ", cls.cCtaCaracter ", "") _
         & "  FROM OpeCta as a JOIN CtaCont as b ON b.cCtaContCod = a.cCtaContCod " _
         & IIf(pbGetClase, " JOIN CtaContClase cls ON a.cCtaContCod LIKE cls.cCtaContCod + '%' ", "") _
         & " where a.cOpeCod = '" & psOpeCod & "' " & lsFiltroDH & lsFiltroOrden _
         & " order by a.cOpeCtaDH"
      Set CargaOpeCta = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeCtaErr:
   MsgBox Err.Description, vbCritical
End Function

Public Function CargaOpeDoc(psOpeCod As String, Optional psMetodo As String, Optional LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
   On Error GoTo CargaOpeDocErr
   Set dbConec = New COMConecta.DCOMConecta
   Dim lsFiltroMet As String
   If psMetodo <> "" Then
        lsFiltroMet = " AND a.cOpeDocMetodo ='" & psMetodo & "'"
   End If
   If dbConec.AbreConexion() Then
      psSql = "SELECT a.cOpeCod, a.nDocTpo, d.cDocDesc, " & _
       "CASE WHEN a.cOpeDocEstado = '11'  THEN 'Obligatorio - Debe Existir' " & _
       "     WHEN a.cOpeDocEstado = '12'  THEN 'Obligatorio - No debe Existir' " & _
       "     WHEN a.cOpeDocEstado = '21'  THEN 'Opcional - Debe Existir' " & _
       "     WHEN a.cOpeDocEstado = '22'  THEN 'Opcional - No debe Existir' END as Estado, " & _
       "CASE WHEN a.cOpeDocMetodo = '1' THEN 'Nro. de Movimiento' " & _
       "     WHEN a.cOpeDocMetodo = '2' THEN 'Autogenerado' " & _
       "     WHEN a.cOpeDocMetodo = '3' THEN 'Digitado' END as Metodo , D.cDocAbrev " & _
       "FROM " & sCentralCom & "OpeDoc a JOIN " & sCentralCom & "Documento d ON d.nDocTpo = a.nDocTpo " & _
       "WHERE cOpeCod = '" & psOpeCod & "' " & lsFiltroMet
       
      Set CargaOpeDoc = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeDocErr:
  MsgBox Err.Description, vbCritical
End Function

Public Function CargaOpeDocEstado(psOpeCod As String, Optional psMetodo As String, Optional psOpeEstado As String, Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
   On Error GoTo CargaOpeDocErr
   Set dbConec = New COMConecta.DCOMConecta
   Dim lsFiltroMet As String
   Dim lsFiltroEstado As String
   If psMetodo <> "" Then
        lsFiltroMet = " AND a.cOpeDocMetodo ='" & psMetodo & "'"
   End If
   If psOpeEstado <> "" Then
        lsFiltroEstado = " AND a.cOpeDocEstado like '" & psOpeEstado & "'"
   End If
   
   If dbConec.AbreConexion() Then
      psSql = " SELECT a.cOpeCod, a.nDocTpo, d.cDocDesc, " & _
       " CASE WHEN a.cOpeDocEstado = '11'  THEN 'Obligatorio - Debe Existir' " & _
       "      WHEN a.cOpeDocEstado = '12'  THEN 'Obligatorio - No debe Existir' " & _
       "      WHEN a.cOpeDocEstado = '21'  THEN 'Opcional - Debe Existir' " & _
       "      WHEN a.cOpeDocEstado = '22'  THEN 'Opcional - No debe Existir' END as Estado, " & _
       " CASE WHEN a.cOpeDocMetodo = '1' THEN 'Nro. de Movimiento' " & _
       "      WHEN a.cOpeDocMetodo = '2' THEN 'Autogenerado' " & _
       "      WHEN a.cOpeDocMetodo = '3' THEN 'Digitado' END as Metodo , D.cDocAbrev " & _
       " FROM " & sCentralCom & "OpeDoc a JOIN " & sCentralCom & "Documento d ON d.nDocTpo = a.nDocTpo " & _
       " WHERE cOpeCod = '" & psOpeCod & "' " & lsFiltroMet & lsFiltroEstado
       
      Set CargaOpeDocEstado = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeDocErr:
   MsgBox Err.Description, vbCritical
End Function

Private Sub Class_Initialize()
    Dim oImp As COMDConstSistema.DCOMImpresoras
    Set oImp = New COMDConstSistema.DCOMImpresoras
       
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

Dim oIni As New COMConecta.DCOMClasIni
sConexion = oIni.CadenaConexion
sCentralCom = oIni.BaseComunes
Set oIni = Nothing
End Sub
